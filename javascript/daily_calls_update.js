var macro ="" ;
var i, retcode, errtext;
var datasource = "C:\\Users\\rzdziech\\Macros\\Datasources\\";
var rest = "";
var desc = "";
var old = "";
var current_date = "{{!NOW:dd}}/{{!NOW:mm}} - ";
var text_do_zapisania = "To be chased today";
var hylas = 3;
var csv_file_name = hylas + "_desc_open_calls.csv";
var row;
var wrong_queue = new Array("-","Pre<SP>Scheduling","Scheduling","Installation","Post<SP>Installation","Sales","Contract<SP>Termination","Finance","Dispatch","RMA","Accreditation","Deployment","AUP","Hybeam");

//PREPARE AND DOWNLOAD CSV FILES FOR ITERATION

macro = "CODE:";

macro += "URL GOTO=http://oss2.freenoc.lan/OSS/SAP/"+hylas+"/serviceCall\n";
//otwieramy servica call w h1 2ndline ktore nie sa zamkniete

macro += "TAG POS=1 TYPE=A ATTR=TXT:customise<SP>fields\n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-fields-form ATTR=ID:edit-fields CONTENT=%subject:%status:%businessPartner:%createdOn:%description:%queue\n";
macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-fields-form ATTR=ID:edit-0\n";

macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=%Status\n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Equality CONTENT=%!=\n";
macro += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT=Closed\n";
macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit\n";

//'trzebe dodac 2ndline bo inaczej exportuje wszystkie call zamiast tylko z 2ndline Dev Error :)
//macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=%Queue\n";
//macro += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT=2nd<SP>line\n";
//macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit\n";

for (var b = 0; b < wrong_queue.length; b++)
{
	macro +="TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=%Queue \n";
	macro +="TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Equality CONTENT=%!= \n";
	macro +="TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT="+wrong_queue[b]+"\n";
	macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit\n";
}

//'zapisujemy liste callsow w csv w datasource zebysmy zaraz mogli z niej sobie z czytac i uaktualnic service call-e
macro += ("ONDOWNLOAD FOLDER=" + datasource + " FILE="+csv_file_name+" WAIT=YES\n");
macro += "TAG POS=1 TYPE=A ATTR=TXT:export<SP>list\n";

//'czekamy 5s zeby csv sie zdarzyl zapisac bo inaczej wywala error 
macro += "WAIT SECONDS=5\n";

retcode = iimPlay(macro);
	
if (retcode < 0) {              // an error has occured
	rerrtext = iimGetLastError();
	alert(errtext);
}
//END OF DOWNLOADING CSV FILE


// COUNT THE NUMBERS OF ROWS IN CSV FOR LOOP
const CRLF = "\r\n";
const LF = "\n";

var lines = new Array();               

var file_i = imns.FIO.openNode(datasource+csv_file_name);
var text = imns.FIO.readTextFile(file_i);         // Read file into one string

// Determine end-of-line marker
var eol = (text.indexOf(CRLF) == -1) ? LF : CRLF;

// Split into lines (number of lines) NUMBER OF LINES IN CSV
lines = text.split(eol);
eol = lines.length;
//iimDisplay(eol);
// FINISH COUNTING EOL = NUmber of LInes 

// RUN LOOP OVER THE SERVICE CALLS
for (var a=2; a < eol ; a++)
{
		iimSet("row",a);
		
		//PREPARE MACRO FOR ITERATION
		macro = "CODE:";
		macro += "SET !DATASOURCE "+csv_file_name+"\n";
		macro += "SET !DATASOURCE_COLUMNS 8\n";
		
		//'set up loop start from second row
		macro += "SET !DATASOURCE_LINE {{row}}\n";
		macro += "URL GOTO=http://oss2.freenoc.lan/OSS/SAP/"+hylas+"/serviceCall/{{!COL1}}\n";
		macro += "TAG POS=1 TYPE=A ATTR=TXT:Edit\n";
		
		//'pobieramy dane z formularza i zmieniamy date lub dodajemy wpis jak brakuje
		macro += "TAG POS=1 TYPE=TEXTAREA FORM=ID:OSSMeta-form ATTR=ID:edit-description EXTRACT=TXT";
		
		//play macro is needed for iimGetLastExtract function to work properly
		iimPlay(macro);
		
		//iimGetLastExtract zwraca "" jezeli jest puste //wrzucamy zawartosc z description do zmiennej
		desc = "";
		old = "";
		rest = "";
		desc = iimGetLastExtract();
		old = iimGetLastExtract();
		var ind = 0;
		var i = 0;
		macro = "CODE:";
		//sprawdzamy co jest w description
		if ( desc == "" || desc == "NULL" || desc == " " ) {
		  // jezeli nic nie ma to dodaj jakis text dzisiajsza date + text 
			desc = current_date + text_do_zapisania + desc;
		}else if ( !isNaN(desc[0]) ){
		 //jezeli pierwszy znak liczba (jakas tam data) to daj nowa i wpisz co bylo poprzednio
			
		 	if ( desc.indexOf("-") > 0 )
		 	{
		 	 	ind = desc.indexOf("-");
		 	 	ind += 1;
		 	}else
		 	{
		 		ind = desc.indexOf(" ");
		 	}
			for (var i = ind; i < desc.length; i++)
			{
			 	rest = rest + desc[i];
			}
			macro +="SET !VAR1 \""+ rest + "\"\n";
			desc = (""+ current_date +"{{!VAR1}}");
			
		} else {
		// jezeli nie mam daty to ja dodaj ja zapisz reszte co bylo wczesniej
			desc = (""+ current_date + "" + desc + "");
		}
		
		//zapisujemy uaktualnione dane
		macro += "TAG POS=1 TYPE=TEXTAREA FORM=ID:OSSMeta-form ATTR=ID:edit-description CONTENT=\"" + desc + "\"\n";
		macro += "TAG POS=1 TYPE=TEXTAREA FORM=ID:OSSMeta-form ATTR=ID:edit-reason CONTENT=\"Last description: " + old + "\"\n";
		macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-form ATTR=ID:edit-0 \n";
		retcode = iimPlay(macro);
		
	if (retcode < 0) {              // an error has occured
	    errtext = iimGetLastError();
	    alert(errtext);
	}

}

