package Spreadsheet::ParseExcel_XLHTML;

$VERSION = 0.01;
@ISA = qw(Spreadsheet::ParseExcel);

=head1 NAME

Spreadsheet::ParseExcel_XLHTML - Parse Excel Spreadsheets using xlhtml

=head1 SYNOPSIS

	use Spreadsheet::ParseExcel_XLHTML;

	my $excel = new Spreadsheet::ParseExcel_XLHTML;

	my $book = $excel->Parse('/some/excel/file.xls');

	for my $sheet (@{$book->{Worksheet}}) {
		print "Worksheet: ", $sheet->{Name}, "\n";
		for (my $i = $sheet->{MinRow}; $i <= $sheet->{MaxRow} ; $i++) {
			print join ',', map { qq|"$_"| }
					map { defined $_ && $_->Value ? $_->Value : "" }
					@{$sheet->{Cells}[$i]};
			print "\n";
		}
	}

=head1 DESCRIPTION

This module follows the interface of the Spreadsheet::ParseExcel module, except
only the "Value" fields of cells are filled, there is no extra fancy stuff. The
reason I wrote it was to have a faster way to parse Excel spreadsheets in Perl.
This module parses around six times faster according to my own informal
benchmarks then the original Spreadsheet::ParseExcel at the time of writing.

To achieve this, it uses a utility called "xlhtml", which you can find here:

	http://www.xlhtml.org/

Get the latest developer release, I've included a patch for 0.3.9.6 that fixes
a couple minor issues, just in case. Don't apply it for a later version.  Once
compiled, it needs to be in the PATH of your Perl program for this module to
work correctly.

You only need to use this module if you have a large volume of big Excel
spreadsheets that you are parsing, otherwise stick to the
Spreadsheet::ParseExcel module.

=cut

use strict;
use IO::File;
use Spreadsheet::ParseExcel;

sub new($;%) {
	my ($class, %args) = @_;
	$class		   = ref $class || $class;

	my $self = bless { %args }, $class;

	return $self;
}

sub Parse($$;$) {
	my ($self, $file) = @_;

	$file = $self->_getFileByObject($file);

	my $stream = new IO::File "xlhtml -xml $file |"
		or die "Could not run xlhtml -xml $file: $!";

	my $work_book = new Spreadsheet::ParseExcel::Workbook;
	$work_book->{File} = $file;

# Start parsing the stream, by hand. An XML parsing method here would be either
# too slow or too complex.
	my @work_sheets;
	my ($sheet, $cells);

	while (<$stream>) {
		chomp;
# Some versions have a bug with the NotImplemented tag getting translated into
# entities...
		s/\&lt;NotImplemented\/\&gt;//;

		/<cell \s* row="(\d+)" \s* col="(\d+)">
		 (?:<[^<>]*>)* # Any <B>, <I> etc. tags
		 ( \s* (?:[^<\s]+\s*?)* )\s* # The value itself.
		 <
		/x && do {
			next if $3 =~ /^\s*$/;

			$cells->[$1][$2] = bless {
				_Value => parseEntities($3)
			}, 'Spreadsheet::ParseExcel::Cell';
			next;
		};
		/<page>(\d+)<\/page>/ && do {
			$work_sheets[$1] = $sheet =
				new Spreadsheet::ParseExcel::Worksheet;
			$sheet->{Cells} = $cells = [];
			next;
		};
		/<pagetitle> (?:<[^<>]*>)* ( \s* (?:[^<\s]+\s*?)* )\s*
		</x && do {
			$sheet->{Name} = parseEntities($1); next;
		};
		/<firstrow>(\d+)<\/firstrow>/ && do {
			$sheet->{MinRow} = $1; next;
		};
		/<lastrow>(\d+)<\/lastrow>/ && do {
			$sheet->{MaxRow} = $1; next;
		};
		/<firstcol>(\d+)<\/firstcol>/ && do {
			$sheet->{MinCol} = $1; next;
		};
		/<lastcol>(\d+)<\/lastcol>/ && do {
			$sheet->{MaxCol} = $1; next;
		};
		/<author>([^<]+)<author>/ && do {
			$work_book->{Author} = $1; next;
		}
	}

	$work_book->{SheetCount} = scalar @work_sheets;
	$work_book->{Worksheet} = \@work_sheets;

	return $work_book;
}

sub parseEntities {
	$_ = shift;
	if (/&/) {
		s/&quot;/"/g; s/&amp;/&/g; s/&gt;/>/; s/&lt;/</;
	}
	return $_;
}

sub DESTROY {
	my $self = shift;

	if (exists $self->{DeleteFiles}) {
		unlink @{$self->{DeleteFiles}};
	}
}

sub _getFileByObject {
	my ($self, $file) = @_;

	if (ref $file eq 'SCALAR') {
		my $file_name = "/tmp/ParseExcel_XLHTML_$$.xls";

		push @{$self->{DeleteFiles}}, $file_name;

		my $writer = new IO::File "> $file_name"
			or die "Could not write to $file_name: $!";

		print $writer $$file;

		close $writer;

		return $file_name;
	} elsif (my $type = ref $file) {
		die "Don't know how to parse file objects of type $type";
	}

	return $file;
}

1;

__END__

=head1 AUTHOR

Rafael Kitover (caelum@debian.org)

=head1 COPYRIGHT

This program is Copyright (c) 2001,2002 by Rafael Kitover. This program is free
software; you can redistribute it and/or modify it under the same terms as Perl
itself.

=head1 ACKNOWLEDGEMENTS

Thanks to the authors of Spreadsheet::ParseExcel and xlhtml for allowing us to
deal with Excel files in the UNIX world.

Thanks to my employer, Gradience, Inc., for allowing me to work on projects
as free software.

=head1 BUGS

Probably a few.

=head1 TODO

I'll take suggestions.

=head1 SEE ALSO

L<Spreadsheet::ParseExcel>,
L<xlhtml>

=cut
