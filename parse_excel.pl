#!/usr/bin/perl 
#use warnings;
use strict;
use Data::Dumper;
use 5.10.1;
use Spreadsheet::ParseXLSX;
use Carp;
use Excel::Writer::XLSX;
use utf8;
use open ':std', ':encoding(UTF-8)';
use Sort::Naturally qw(ncmp);

my $file = $ARGV[0];


#TODO:
# mapping colums to be processed in same functions as general file
# do lc
# reorder mapping to be alphabetic style

my $mapping = read_map( $ARGV[1] );
die Dumper $mapping;

confess "You must give an file name as first parameter" unless $file;

my $result_data = {};
my $result_file = "Result.xlsx";
my $result_sheet_name = 'Result';
my $rows_to_proces = 243;
my $column_to_read = 8;

############
my $parser =   Spreadsheet::ParseXLSX->new();
############

say "\n###################";
say "Parse file[$file]";
say "###################\n";

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

			  if ($col eq $column_to_read ) {
				  $counter++;
				  #say "Row[$row] col[$col]"; 
				  my $val = $cell->value();
				  my $old_value = $val;
				  #say "before func[$val]";
				  my $initial_string = $val;
				  my ($uni_count, $abbr_count );

				  $val = trim($val);
				  $val = rem_spaces($val);
				  #$val = case($val);

				  ($val, $abbr_count ) = abriviation2($val);
				  #($val, $abbr_count ) = abriviation($val);
				  ($val, $uni_count ) = unification($val);
				
				  say "[$initial_string] ---> [$val]";

				  $result_data->{$row}{new_value} = $val;
				  $result_data->{$row}{old_value} = $old_value;
				  $result_data->{$row}{row_num} = $row + 1;
				  $result_data->{$row}{col_num} = 2;
				  $result_data->{$row}{abbr} = $abbr_count;
				  $result_data->{$row}{uni} = $uni_count;

			  }

			  if ($col eq 1 ) {
				  my $id = $cell->value();
				  $result_data->{$row}{id} = $id;

			  }
	  		  last ROW_NUM if $counter > $rows_to_proces;
		  }
	  }
}
say "\nAbout to write data in file[$result_file] with sheet[$result_sheet_name]\n";
write_down( $result_file, $result_sheet_name , $result_data );

################
sub write_down { 
################
	my $file = shift;
	my $sheet_name = shift;
	my $new_excel_data = shift;
	#say "show entire data" . Dumper $new_excel_data;	

	my $workbook = Excel::Writer::XLSX->new( $file );

	# Add a worksheet
	my $worksheet = $workbook->add_worksheet($sheet_name);

	#write headers
	$worksheet->write( 0 , 0 , "ID" );
	$worksheet->write( 0 , 1 , "OLD_VALUE" );
	$worksheet->write( 0 , 2 , "NEW_VALUE" );
	$worksheet->write( 0 , 3 , "abbr" );
	$worksheet->write( 0 , 4 , "uni" );

	for my $row_in ( %{$new_excel_data} ) {
		next if ref $row_in; 	
		my $row = $new_excel_data->{$row_in}{row_num};
		my $col = $new_excel_data->{$row_in}{col_num};
		my $new_value = $new_excel_data->{$row_in}{new_value};
		my $id = $new_excel_data->{$row_in}{id};
		my $old_value = $new_excel_data->{$row_in}{old_value};
		my $abbr_count = $new_excel_data->{$row_in}{abbr};
		my $uni_count = $new_excel_data->{$row_in}{uni};
	#say "In write_down() row_in[$row_in] row[$row] col[$col] new_value[$new_value]";

		#ID
		$worksheet->write( $row , 0, $id );
		#OLD_value
		$worksheet->write( $row, 1, $old_value );
		#nEW_value	
		$worksheet->write( $row, $col, $new_value );
		#abbr
		$worksheet->write( $row, $col + 1, $abbr_count );
		#uni
		$worksheet->write( $row, $col + 2 , $uni_count );
	}	
	$workbook->close();
}
##########
sub trim {
##########
	my $value = shift;
	$value =~ s/^\s*//;
	$value =~ s/\s*$//;
	return $value;
}
################
sub rem_spaces {
################
	my $value = shift;
	$value =~ s/\s\s*/ /g; 
	return $value;
}
##########
sub case {
##########
	my $value = shift;
	my @words = split ' ', $value;
	my $result;
	for my $word ( @words) {

		$result .= ucfirst( lc($word) ) . ' ';
	}	
	return trim($result);


}



#################
sub abriviation2 {
#################
	my $input_value  = shift;
	my $matched;
	my $abbrs = {
		'1' => {	'- ' => '-',},
		'2' => {	' -' => '-',},
		'3' => {	' - ' => '-',},
		'4' => {	'-Т' => '-т',},
		'5' => {	'-Р' => '-р',},
		'6' => {	'-В' => '-в',},
		'7' => {	'-М' => '-м',},
		'8' => {	' И ' => ' и ',},
		'9' => {	'Ii' => 'II',},
		'10' => {	'Iii' => 'III',},
		'11' => {	'Академик' => 'акад.',},
		'12' => {	'Архитект' => 'арх.',},
		'13' => {	'Генерал' => 'ген.',},
		'14' => {	'Доктор' => 'д-р',},
		'15' => {	'Доцент' => 'доц.',},
		'16' => {	'Инженер' => 'инж.',},
		'17' => {	'Капитан' => 'к-н.',},
		'18' => {	'Майор' => 'м-р',},
		'19' => {	'Лейтенант' => 'лейт.',},
		'20' => {	'Подполковник' => 'подпл.',},
		'21' => {	'Подпоручик' => 'подпр.',},
		'22' => {	'Полковник' => 'полк.',},
		'23' => {	'Поручик' => 'прк.',},
		'24' => {	'Професор' => 'проф.',},
		'25' => {	'Акад.' => 'акад.',},
		'26' => {	'Арх.' => 'арх.',},
		'27' => {	'Ген.' => 'ген.',},
		'28' => {	'Доц.' => 'доц.',},
		'29' => {	'Инж.' => 'инж.',},
		'30' => {	'Кап.' => 'к-н.',},
		'32' => {	'М-Р' => 'м-р',},
		'33' => {	'Лейт.' => 'лейт.',},
		'34' => {	'Подпл.' => 'подпл.',},
		'35' => {	'Подпр.' => 'подпр.',},
		'36' => {	'Полк.' => 'полк.',},
		'37' => {	'Прк.' => 'прк.',},
		'38' => {	'Проф.' => 'проф.',},
		'39' => {	'Д-Р' => 'д-р',},
	};
	
	for my $abbr_key ( sort {  ncmp( $a,  $b) } keys %{$abbrs} ) {	
		my @key_value =  keys %{$abbrs->{$abbr_key}};
		my $map_key = $key_value[0];
		my $map_result = $abbrs->{$abbr_key}->{$map_key};
		#say "abr_key[$abbr_key] key[$map_key] map_result[$map_result]";

		if ($input_value =~ /\Q$map_key\E/i ) {
			$matched++;
		}
		$input_value =~ s/\Q$map_key\E/$map_result/gi;
		
	}

	#say "result of abriviation() [$input_value] matched[$matched]";
	return $input_value, $matched;
}

#################
sub abriviation {
#################
	my $input_value  = shift;
	my $matched;
	my $abbrs = {
		'- ' => '-',
		' -' => '-',
		' - ' => '-',
		'-Т' => '-т',
		'-Р' => '-р',
		'-В' => '-в',
		'-М' => '-м',
		' И ' => ' и ',
		'Ii' => 'II',
		'Iii' => 'III',
		'Академик' => 'акад.',
		'Архитект' => 'арх.',
		'Генерал' => 'ген.',
		'Доктор' => 'д-р',
		'Доцент' => 'доц.',
		'Инженер' => 'инж.',
		'Капитан' => 'к-н.',
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
		'Кап.' => 'к-н.',
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
		if ($input_value =~ /\Q$abbr_key\E/i ) {
			$matched++;
		}
		$input_value =~ s/\Q$abbr_key\E/$abbr_value/gi;
	}
	#say "result of abriviation() [$input_value] matched[$matched]";
	return $input_value, $matched;
}
#################
sub unification {
#################
	my $input_value = shift;
	my $matched;
	my $uni_list  = {
		'Алек.*?Стамб[^\s]*' => 'Александър Стамболийски',
		'Нико.*?Миха[^\s]*' => 'Никола Михайловски',
		'първа(\s|$)' => '1-ва',
		'втора(\s|$)' => '2-ра',
		'трета(\s|$)' => '3-та',
		'1(\s|$)' => '1-ва',
	};

	for my $abbr_key ( keys %{$uni_list} ) {	
		my $abbr_value = $uni_list->{$abbr_key};
		#say "input_value[$input_value] about to replase[$abbr_key] to [$abbr_value]";
		if ( $input_value =~ /$abbr_key/i ) {
			$matched++;
		}

		$input_value =~ s/$abbr_key/$abbr_value/i;
	}
	return $input_value, $matched;
}


###################
sub read_map {
###################
	my $excel_file = shift;
	my $result_data;

	my $parser =   Spreadsheet::ParseXLSX->new();
	my $workbook = $parser->parse($excel_file);

	for my $worksheet ( $workbook->worksheets() ) {
		my ( $row_min, $row_max ) = $worksheet->row_range();
		my ( $col_min, $col_max ) = $worksheet->col_range();

ROW_NUM:  for my $row ( $row_min .. $row_max ) {
		  my $result_key;
		  my $result_value;
		  for my $col ( $col_min .. $col_max ) {
			  my $cell = $worksheet->get_cell( $row, $col );


			  my $cell_value;
			  if ($cell) {

				  $cell_value  = $cell->value();
			  } else {
				  $cell_value = '';
			  }		

			  if ($col eq '0' ) {
				  $result_key = $cell_value;
			  }


			  if ($col eq '1' ) {
				  $result_value = $cell_value;
			  }


		  }
		  $result_data->{$row} = { $result_key => $result_value };
		  #say "column_read[\$col] row[$row] key[$result_key] value[$result_value] ";
	  }
	}

	return $result_data;

}


