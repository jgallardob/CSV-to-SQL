#Script extract information from ODS file, and generate a SQL file
#Created by Juan Gallardo Beltran
#jgallarodb2014@alu.uct.cl
#Information Managment - INFO1107
#Professor Guide Alejandro Mellado
#Temuco, Chile - May 2016

#!/usr/bin/perl

#Libraries to use
use OpenOffice::OODoc;
use strict;
use Data::Dumper;
#Open de ODS File
my $doc = odfDocument(file=> 'lista_de_usuarios(1).ods');
#Generate the head of SQL FILE
my $ini = "use DATABASE registros;";
my $in2 = "\ninsert into nomina (rut,dv,registro,ap_pa,ap_ma,nombres,sexo,fecha_nacto,estado,ano_ingreso,cod_carr,nom_carr,plan,mencion) VALUES";
#Make the XML file, and add the head of XML
open(SALIDA,">>file.sql");
print SALIDA $ini,$in2;
#Take the first position and extract the data into the XML file
	for (my $i=6;$i<328;$i++){
	my $rut = $doc -> getCellValue(0, "A$i");
	my $dv  = $doc -> getCellValue(0, "B$i");
	my $reg = $doc -> getCellValue(0, "C$i");
	my $apa = $doc -> getCellValue(0, "D$i");
	my $ama = $doc -> getCellValue(0, "E$i");
	my $nom = $doc -> getCellValue(0, "F$i");
	my $sex = $doc -> getCellValue(0, "G$i");
	my $nct = $doc -> getCellValue(0, "H$i");
	my $sta = $doc -> getCellValue(0, "I$i");
	my $ing = $doc -> getCellValue(0, "J$i");
	my $cod = $doc -> getCellValue(0, "K$i");
	my $nca = $doc -> getCellValue(0, "L$i");
	my $pla = $doc -> getCellValue(0, "M$i");
	my $men = $doc -> getCellValue(0, "N$i");
	#Here, write the records into XML file
	my $file = "($rut, $dv, $reg, $apa, $ama, $nom, $sex, $nct, $sta, $ing, $cod, $nca, $pla, $men);";
	print SALIDA $file;
}
#Add de GO order
my $fin = "\nGO	";
#Write the record into a SQL file
print SALIDA $fin;
#Close the files
close SALIDA;
