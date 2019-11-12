#!/usr/bin/perl 
#use warnings;
use strict;
use Data::Dumper;
use 5.10.1;
use Spreadsheet::ParseXLSX;
use Carp;
use Excel::Writer::XLSX;


my $file = $ARGV[0];


confess "You must give an file name as first parameter" unless $file;
my $result_data = {};
# my $result->{1} = {
#	          old_value =>'iasdasd', 
#		  new_value => "asdasd123123" ,
#		  id => '123123', 
#		  row_num => '1',
#		  col_num => '2' #HARD CODED
#		  #ID col 0, 
#		  #OLD_VALUE col 1
#                 };

my $result_file = "Result.xlsx";

##############

my $parser =   Spreadsheet::ParseXLSX->new();
#my $parser   = Spreadsheet::ParseExcel->new();
say "Parse file[$file]";
my $workbook = $parser->parse($file);
 
if ( !defined $workbook ) {
    confess $parser->error() .  ".\n";
}


my $counter;


for my $worksheet ( $workbook->worksheets() ) {
	say "show me worksheet name[$worksheet]"; 
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
				#$val = abriviation($val);
				say "after[$val]";
				#say "Unformatted["  .  $cell->unformatted() . "]";
                                say "\n";
				
                                $result_data->{$counter}{new_value} = $val;

                                $result_data->{$counter}{old_value} = $old_value;
                                $result_data->{$counter}{row_num} = $row + 1;
                                $result_data->{$counter}{col_num} = 2;
				
				
			
			}

			if ($col eq 1 ) {
				my $id = $cell->value();
                                $result_data->{$counter}{id} = $id;
                           
			}
			last ROW_NUM if $counter > 10;
		}
	}
}


write_down( $result_file, 'Result' , $result_data );

######################
sub transliterate {
######################
      my $value = shift;
      	

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
# академик Иван Гeшoв , акдемик Иван Ге
	my $abbrs = {

		'академик' => 'акад.',
		'генерал' => 'ген.',
	};

	for my $abbr_key ( keys %{$abbrs} ) {	
               my $abbr_value = $abbrs->{$abbr_key};
	       $input_value =~ s/$input_value/$abbr_value/g;
        }

   return $input_value;

}

#####################
sub unification {
#####################
	my $input_value = shift;
	my $uni_list  = {
		#Алекдандер Стамболиики 
		qr/Алек.*?Стамб/ => 1,
        };

	for my $uni_key ( keys %{$uni_list} ){

		if ( $input_value =~ /$uni_key/i ){
		     $input_value =~ s/$input_value/$uni_key/g;
		     last;
		}
	}
	
	return $input_value;

}

########################
sub write_down { 
########################
	my $file = shift;
	my $sheet_name = shift;
	my $new_excel_data = shift;
	say "show entire data" . Dumper $new_excel_data;	
	

	my $workbook = Excel::Writer::XLSX->new( $file );

        # Add a worksheet
	my $worksheet = $workbook->add_worksheet($sheet_name);

        #  Add and define a format
	#$format = $workbook->add_format();
	#$format->set_bold();
	#$format->set_color( 'red' );
	#$format->set_align( 'center' );

        # Write a formatted and unformatted string, row and column notation.
#my $result = {  1 => {
#	          old_value =>'iasdasd', 
#		  new_value => "asdasd123123" ,
#		  id => '123123', 
#		  row_num => '1',
#		  col_num => '2'
#                 }, 
#             };

		
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
		say "In write_down() row_in[$row_in] row[$row] col[$col] new_value[$new_value]";
	
		#ID
		$worksheet->write( $row, 0, $id );
		#OLD_value
		$worksheet->write( $row, 1, $old_value );
		#nEW_value	
		$worksheet->write( $row, $col, $new_value );
	}	

        # Write a number and a formula using A1 notation
	#$worksheet->write( 'A3', 1.2345 );
	#$worksheet->write( 'A4', '=SIN(PI()/4)' );

	$workbook->close();
}



