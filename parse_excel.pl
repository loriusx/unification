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
use Getopt::Long qw(GetOptionsFromArray);
Getopt::Long::Configure qw(ignorecase_always permute);
use POSIX 'strftime';



my $data   = "file.dat";
my $length = 24;
my $verbose;

#TODO:
# Fast main excel file ( skip useless cols )
# fork if is too slow
# uncomment case() maybe
# make id column to be in global variable


my $start_time = &time_;
my $Options = _makeOptionsFromArray( @ARGV );

my $file = $Options->{data_file};
my $criteria_file = $Options->{criteria_file};
my $rows_to_proces = $Options->{rows_to_proces};


for my $manda_arg ( qw(data_file criteria_file ) ) {
	unless ($Options->{$manda_arg} ) {
		confess "Mandatory argument[$manda_arg] is missing for $0";
	}

}


my $mapping = read_map( $criteria_file );

my $result_data = {};
my $result_file = $Options->{result_file} || "Result.xlsx";
my $result_sheet_name = 'Result';
my $column_to_read = 3;

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
				  my ($uni_count, $abbr_count );

				  my $func_result = func_calls( $val );
				  $val = $func_result->{value};

	                          my $crit_val = do_criteria($val, $mapping);
					
                                  if ($crit_val) {
				     $val = $crit_val;				  
				  }

				  say "[$old_value] ---> [$val]";

				  $result_data->{$row}{new_value} = $val;
				  $result_data->{$row}{old_value} = $old_value;
				  $result_data->{$row}{row_num} = $row + 1;
				  $result_data->{$row}{col_num} = 2;
				  $result_data->{$row}{abbr} = $func_result->{abbr_count};
				  $result_data->{$row}{uni} = $func_result->{uni_count};

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

my $end_time = &time_;

say "start script time[$start_time] end time[$end_time]";

##################
sub func_calls {
##################
	my $val = shift;
	my ($abbr_count, $uni_count );
	my $result = {};

	$val = lc( $val );
	$val = trim($val);
	$val = rem_spaces($val);
	$val =~ s/\./\. /g;
        $val = case($val);

	($val, $abbr_count ) = abriviation($val);

	$result->{value} = $val;
	$result->{uni_count} = $uni_count;
	$result->{abbr_count} = $abbr_count;

	return $result;

}

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
		$worksheet->write( $row, 0, $id );
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
sub abriviation {
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
		'17' => {	'Капитан' => 'к-н',},
		'18' => {	'Майор' => 'м-р',},
		'19' => {	'Мичман' => 'м-н',},		
		'20' => {	'Лейтенант' => 'лейт.',},
		'21' => {	'Подполковник' => 'подпл.',},
		'22' => {	'Подпоручик' => 'подпр.',},
		'23' => {	'Полковник' => 'полк.',},
		'24' => {	'Поручик' => 'прк.',},
		'25' => {	'Професор' => 'проф.',},
		'26' => {	'Акад.' => 'акад.',},
		'27' => {	'Арх.' => 'арх.',},
		'28' => {	'Ген.' => 'ген.',},
		'29' => {	'Доц.' => 'доц.',},
		'30' => {	'Инж.' => 'инж.',},
		'31' => {	'Кап.' => 'к-н',},
		'32' => {	'М-Р' => 'м-р',},
		'33' => {	'Лейт.' => 'лейт.',},
		'34' => {	'Подп.' => 'подп.',},
		'35' => {	'Подпр.' => 'подпр.',},
		'36' => {	'Полк.' => 'полк.',},
		'37' => {	'Прк.' => 'прк.',},
		'38' => {	'Проф.' => 'проф.',},
		'39' => {	'Д-Р' => 'д-р',},
		'40' => {	'Цар ' => 'цар ',},
		'41' => {	'Хан ' => 'хан ',},
		'42' => {	'Пл.' => 'пл.',},
		'43' => {	'Княз' => 'княз',},
		'44' => {	'Епископ' => 'епископ',},
		'45' => {	'Архиепископ' => 'архиепископ',},
		'46' => {	'Св. ' => 'св. ',},
		'47' => {	'Отец ' => 'отец ',},
		'48' => {	'Йеромонах ' => 'йеромонах ',},
		'49' => {	'Иконом ' => 'иконом ',},
		'50' => {	'Опълченец ' => 'опълченец ',},
		'51' =>	{	'Митрополит ' => 'митрополит ',},
		'52' => {	'Свещеник ' => 'свещеник ',},
		'53' => {	'Княгиня ' => 'княгиня ',},
		'54' => {	'Алея ' => 'алея ',},
		'55' => {	'Архимандрит ' => 'архимандрит ',},
		'56' => {	'Лесничей ' => 'лесничей ',},
		'57' => {	'Поп ' => 'поп ',},
		'58' => {	'Папа ' => 'папа ',},
		'59' => {	'Парк ' => 'парк ',},
		'60' => {	'Пастор ' => 'пастор ',},
		'61' => {	'Патриарх ' => 'патриарх ',},
		'62' => {	'Подофицер ' => 'подофицер ',},
		'63' => {	'Воевода ' => 'Войвода ',},
		'64' => {	' Княгиня' => ' княгиня',},
		'65' => {	' Дивизия' => ' дивизия',},
		'66' => {	' Полк' => ' полк',},
		'67' => {	' Ливада' => ' ливада',},
		'68' => {	' войвода' => ' Войвода',},
		'69' => {	' Воевода' => ' Войвода',},
		'70' => {	' Километър' => ' километър',},
		'71' => {	' Конгрес' => ' конгрес',},
		'72' => {	' Долина' => ' долина',},
		'73' => { 	' Планина' => ' планина',},
		'74' => {	' Извор' => ' извор',},
		'75' => {	' Река' => ' река',},
		'76' => {	' Шосе' => ' шосе',},

	};	
	for my $abbr_key ( sort {  ncmp( $a,  $b) } keys %{$abbrs} ) {	
		#say "show me abbr_key[$abbr_key]";
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
#NOTE sort may slow the process
#
#################
sub do_criteria {
#################
	my $input_value = shift;
	my $mapping = shift;
	
	my $new_value;

	#for my $row_num ( sort {  ncmp( $a,  $b) } keys %{$mapping} ) {
	for my $row_num (  keys %{$mapping} ) {

		my @key_value =  keys %{$mapping->{$row_num}};
		my $old_value = $key_value[0];
		my $to_be = $mapping->{$row_num}->{$old_value};
		next unless $to_be;
	#	say "input_Value[$input_value] will be compared with [$old_value]"; 

		if ( $input_value eq $old_value ) {
			$new_value = $to_be;
			say "Input_value[$input_value] has matched with old_value[$old_value] and will became[$to_be]";
			last;
		}

	}

	return $new_value
}


###################
sub read_map {
###################
	my $excel_file = shift;
	my $result_data;

	say "#################################";
	say "Parsing file mapping[$excel_file]";
	say "#################################";
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
				  my $result_call = func_calls($cell_value);
				  $result_key = $result_call->{value};
			  }

			  if ($col eq '1' ) {
				  $result_value = $cell_value;
			  }


		  }
		  $result_data->{$row} = { $result_key => $result_value };
		  #say "row[$row] key[$result_key] value[$result_value] ";
	  }
	}

	return $result_data;

}

##################################
sub _makeOptionsFromArray {
##################################
    my @Options_array = @_;
    my $Options       = {};

    if ($Options_array[0] and $Options_array[0] !~ /^--?/ ){ #prevent user mistake
        confess "Error: Wrong argv. First \@ARGV argumet must start with single or double dash -|-- false arg[$Options_array[0]";
    }

    if (!@Options_array) {
	    print "$0: Argument required --data-file <file_path> and --criteria-file <file_path>.\n";
	    exit 1;
    }

    GetOptionsFromArray(
        \@Options_array,
		"data-file=s"   => \$Options->{data_file}, 
		"criteria-file=s"   => \$Options->{criteria_file},
		"rows-process=i"    => \$Options->{rows_to_proces},
                "result-file=s"    => \$Options->{result_file},
		"help"    => \$Options->{help},

    ) or confess( "Err: command line arguments are wrong\n" );

    if ( $Options->{ help } ) {
        print "$0: Argument required --data-file <file_path> and --criteria-file <file_path>.\n";
        exit 0;
    }

    return $Options;
}

################
sub time_ {
################
	my $time = strftime('%T', localtime);
	return $time;
}
