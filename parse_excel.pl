#!/usr/bin/perl 
#use warnings;
use strict;
use Data::Dumper;
use 5.10.1;
use Spreadsheet::ParseXLSX;
use Carp;
use Excel::Writer::XLSX;
use utf8;

my $file = $ARGV[0];

confess "You must give an file name as first parameter" unless $file;

my $result_data = {};
my $result_file = "Result.xlsx";
my $result_sheet_name = 'Result';
my $rows_to_proces = 10;

##############

my $parser =   Spreadsheet::ParseXLSX->new();

say "Parse file[$file]";

my $workbook = $parser->parse($file);
 
if ( !defined $workbook ) {
    confess $parser->error() .  ".\n";
}

my $counter = 0;


for my $worksheet ( $workbook->worksheets() ) {
	my ( $row_min, $row_max ) = $worksheet->row_range();
	my ( $col_min, $col_max ) = $worksheet->col_range();

ROW_NUM:  for my $row ( $row_min .. $row_max ) {
		  for my $col ( $col_min .. $col_max ) {

			  my $cell = $worksheet->get_cell( $row, $col );
			  next unless $cell;

			  if ($col eq 8 ) {
				  $counter++;
				  say "Row[$row] col[$col]"; 
				  my $val = $cell->value();
				  my $old_value = $val;
				  say "before func[$val]";

				  $val = trim($val);
				  $val = rem_spaces($val);
				  $val = case($val);
				  $val = abriviation($val);
				  $val = unification($val);

				  say "after[$val]";
				  say "\n";

				  $result_data->{$row}{new_value} = $val;
				  $result_data->{$row}{old_value} = $old_value;
				  $result_data->{$row}{row_num} = $row + 1;
				  $result_data->{$row}{col_num} = 2;



			  }

			  if ($col eq 1 ) {
				  my $id = $cell->value();
				  $result_data->{$row}{id} = $id;

			  }
			  last ROW_NUM if $counter > $rows_to_proces;
		  }
	  }
}


write_down( $result_file, $result_sheet_name , $result_data );

########################
sub write_down { 
########################
	my $file = shift;
	my $sheet_name = shift;
	my $new_excel_data = shift;
	#say "show entire data" . Dumper $new_excel_data;	


	my $workbook = Excel::Writer::XLSX->new( $file );

	# Add a worksheet
	my $worksheet = $workbook->add_worksheet($sheet_name);

	#Make headers
	$worksheet->write( 0 , 0 , "ID" );
	$worksheet->write( 0 , 1 , "OLD_VALUE" );
	$worksheet->write( 0 , 2 , "NEW_VALUE" );

	for my $row_in ( %{$new_excel_data} ) {
		next if ref $row_in; 	
		my $row = $new_excel_data->{$row_in}{row_num};
		my $col = $new_excel_data->{$row_in}{col_num};
		my $new_value = $new_excel_data->{$row_in}{new_value};
		my $id = $new_excel_data->{$row_in}{id};
		my $old_value = $new_excel_data->{$row_in}{old_value};
	#say "In write_down() row_in[$row_in] row[$row] col[$col] new_value[$new_value]";

		#ID
		$worksheet->write( $row , 0, $id );
		#OLD_value
		$worksheet->write( $row, 1, $old_value );
		#nEW_value	
		$worksheet->write( $row, $col, $new_value );
	}	

	$workbook->close();
}

################
sub trim {
	my $value = shift;
	$value =~ s/^\s*//;
	$value =~ s/\s*$//;
	return $value;
}

###############
sub rem_spaces {
#############
	my $value = shift;
	$value =~ s/\s\s*/ /g; 
	return $value;

}

#############
sub case {
#############
	my $value = shift;
	my @words = split ' ', $value;
	my $result;

	for my $word ( @words) {

		$result .= ucfirst( lc($word) ) . ' ';
	}	

	return trim($result);

}


####################
sub abriviation {
####################
	my $input_value  = shift;

	my $abbrs = {

		' - ' => '-',
		'-Т' => '-т',
		'-Р' => '-р',
		'-В' => '-в',
		'-М' => '-м',
		' И ' => ' и',
		'Ii' => 'II',
		'Iii' => 'III',
		'Академик' => 'акад.',
		'Архитект' => 'арх.',
		'Генерал' => 'ген.',
		'Доктор' => 'д-р',
		'Доцент' => 'доц.',
		'Инженер' => 'инж.',
		'Капитан' => 'кап.',
		'Майор' => 'м-р',
		'Лейтенант' => 'лейт.',
		'Подполковник' => 'подпл.',
		'Подпоручик' => 'подпр.',
		'Полковник' => 'полк.',
		'Поручик' => 'прк.',
		'Професор' => 'проф.',
		'Акад.' => 'акад.',
		'Арх.' => 'арх.',
		'Ген.' => 'ген.',
		'Доц.' => 'доц.',
		'Инж.' => 'инж.',
		'Кап.' => 'кап.',
		'М-Р' => 'м-р',
		'Лейт.' => 'лейт.',
		'Подпл.' => 'подпл.',
		'Подпр.' => 'подпр.',
		'Полк.' => 'полк.',
		'Прк.' => 'прк.',
		'Проф.' => 'проф.',
		'Д-Р' => 'д-р',

	};

	for my $abbr_key ( keys %{$abbrs} ) {	
		my $abbr_value = $abbrs->{$abbr_key};
		#say "about to replase[$abbr_key] to [$abbr_value]";
		$input_value =~ s/\Q$abbr_key\E/$abbr_value/gi;
	}

	#say "result of abriviation() [$input_value]";
	return $input_value;

}

#####################
sub unification {
#####################
	my $input_value = shift;
	my $uni_list  = {
		'Алек.*?Стамб[^\s]*' => 'Александър Стамболийски',
		'Нико.*?Миха[^\s]*' => 'Никола Михайловски',
		'първа(\s|$)' => '1-ва',
		'втора(\s|$)' => '2-ра',
		'трета(\s|$)' => '3-та',
	};


	for my $abbr_key ( keys %{$uni_list} ) {	
		my $abbr_value = $uni_list->{$abbr_key};
		#say "input_value[$input_value] about to replase[$abbr_key] to [$abbr_value]";
		$input_value =~ s/$abbr_key/$abbr_value/i;
	}

	return $input_value;

}




