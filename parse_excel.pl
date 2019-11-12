#!/usr/bin/perl 
#use warnings;
use strict;
use Data::Dumper;
use 5.10.1;
use Spreadsheet::ParseXLSX;
use Carp;


my $file = $ARGV[0];

my $parser =   Spreadsheet::ParseXLSX->new();
#my $parser   = Spreadsheet::ParseExcel->new();
say "Parse file[$file]";
my $workbook = $parser->parse($file);
 
if ( !defined $workbook ) {
    confess $parser->error() .  ".\n";
}


my $counter;
my $result = { old_value =>'', new_value => "" , id => '' };

for my $worksheet ( $workbook->worksheets() ) {
	say "show me worksheet name[$worksheet]"; 
	my ( $row_min, $row_max ) = $worksheet->row_range();
	my ( $col_min, $col_max ) = $worksheet->col_range();

	for my $row ( $row_min .. $row_max ) {
		for my $col ( $col_min .. $col_max ) {

			my $cell = $worksheet->get_cell( $row, $col );
			next unless $cell;
				
			if ($col eq 8 ) {
				$counter++;
				say "Row[$row] col[$col]"; 
				my $val = $cell->value();
				say "before func[$val]";
                                $val = trim($val);
				$val = rem_spaces($val);
				$val = case($val);
				say "after[$val]";
				#say "Unformatted["  .  $cell->unformatted() . "]";
                                say "\n";
				
				
				die if $counter > 10;
			
			}
		}
	}
}


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

	    $result .= ucfirst( $word ) . ' ';
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


#use Excel::Writer::XLSX;
 
## Create a new Excel workbook
#my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );
# 
## Add a worksheet
#$worksheet = $workbook->add_worksheet();
# 
##  Add and define a format
#$format = $workbook->add_format();
#$format->set_bold();
#$format->set_color( 'red' );
#$format->set_align( 'center' );
# 
## Write a formatted and unformatted string, row and column notation.
#$col = $row = 0;
#$worksheet->write( $row, $col, 'Hi Excel!', $format );
#$worksheet->write( 1, $col, 'Hi Excel!' );
# 
## Write a number and a formula using A1 notation
#$worksheet->write( 'A3', 1.2345 );
#$worksheet->write( 'A4', '=SIN(PI()/4)' );
# 
#$workbook->close();



