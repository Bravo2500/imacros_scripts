// month set up
var macro = "";
var retcode;
var errtext;
var date = new Date();
var start_date = "";
var end_date = "";
var year = date.getFullYear();
var month = date.getMonth() + 1;
var hylas ="";
var last_month = 0;
var datasource_h1 = "h1.csv";
var datasource_h2 = "h2.csv";
var datasource = "C:\\Users\\rzdziech\\Macros\\Datasources\\";

if (month == 1)
{
    
   last_month = 12;

}else{

   last_month = month - 1;

}

var download_folder = "ONDOWNLOAD FOLDER=C:\\Users\\rzdziech\\Desktop\\PR<SP>Automation<SP>Project\\datasource\\imacro_export\\{{!NOW:yyyy}}_"+last_month+" FILE={{!COL2}}.csv WAIT=YES \n";

start_date = ""+year+""+last_month+"01";
end_date = ""+year+""+last_month+"31";
    
macro = "CODE:";
macro += "SET !VAR1 \""+start_date+"\"\n";
macro += "SET !VAR2 \""+end_date+"\"\n";
macro += "TAB CLOSEALLOTHERS \n";
macro += "TAB OPEN \n";
macro += "TAB T=1 \n";
macro += "URL GOTO=http://oss2.freenoc.lan/OSS/SAP/3/serviceCall \n";

//form setting up 
//'TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-reset

//prepare for export

macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=\%Created<SP>Date\/Time \n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Equality CONTENT=%> \n";
macro += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT={{!VAR1}} \n";
macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit \n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=\%Created<SP>Date\/Time \n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Equality CONTENT=%< \n";
macro += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT={{!VAR2}} \n";
macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit \n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=%Subject \n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Equality CONTENT=%!= \n";
macro += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT=AUP<SP>Enforcement \n";
macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit \n";
macro += "TAG POS=1 TYPE=A ATTR=TXT:customise<SP>fields \n";
macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-fields-form ATTR=ID:edit-fields CONTENT=%subject:%status:%businessPartner:%type:%createdOn:%closedOn:%resolution \n";
macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-fields-form ATTR=ID:edit-0 \n";


retcode = iimPlay(macro);
	
if (retcode < 0) {              // an error has occured
	rerrtext = iimGetLastError();
	alert(errtext);
}

// <<<<<<< OSS SETUP HAS FINISHED >>>>>>>>>
//
//count the number of partners in csv
function number_of_lines (datasource, hylas)
{
    const CRLF = "\r\n";
    const LF = "\n";

    var lines = new Array();               

    var file_i = imns.FIO.openNode(datasource+hylas);
    var text = imns.FIO.readTextFile(file_i);         // Read file into one string

    // Determine end-of-line marker
    var eol = (text.indexOf(CRLF) == -1) ? LF : CRLF;

    // Split into lines (number of lines) NUMBER OF LINES IN CSV
    lines = text.split(eol);
    eol = lines.length;
    
    return eol;
}

// <<<<<<< EXPORTING PARTNERS  >>>>>>>>>
// loop over the partners
function loop_over_partners(eol,partner_datasource,download_folder,hylas)
{
        
    // RUN LOOP OVER THE SERVICE CALLS
    for (var a=2; a < eol ; a++)
    {
            iimSet("row",a);

            macro = "CODE:";
            macro +="SET !DATASOURCE "+partner_datasource+"\n";
            macro +="SET !DATASOURCE_COLUMNS 8 \n";
            
            //set up partner CONTENT=SAP_CODE
            //'set up loop start from second row
            macro += "SET !DATASOURCE_LINE {{row}}\n";
                     
            macro += "URL GOTO=http://oss2.freenoc.lan/OSS/SAP/"+hylas+"/serviceCall \n";
            macro += "TAG POS=1 TYPE=SELECT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Column CONTENT=%Business<SP>Partner \n";
            macro += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-Value CONTENT={{!COL2}} \n";
            macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-submit \n";
            
            //set up name for files
            
            macro += "URL GOTO=http://oss2.freenoc.lan/OSS/SAP/"+hylas+"/serviceCall \n";
            macro += ""+download_folder+"\n";
            macro += "TAG POS=1 TYPE=A ATTR=TXT:export<SP>list CONTENT=EVENT:SAVETARGETAS \n";
            
            macro += "WAIT SECONDS=2 \n";
            macro += "TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:OSSMeta-filter-form ATTR=ID:edit-undo \n";
            macro += "WAIT SECONDS=3 \n";
            
            
            retcode = iimPlay(macro);
                
            if (retcode < 0) {              // an error has occured
                rerrtext = iimGetLastError();
                alert(errtext);
            }
    }
}

eol = number_of_lines(datasource,datasource_h1);
loop_over_partners(eol,datasource_h1,download_folder,1)

eol = number_of_lines(datasource,datasource_h2);
loop_over_partners(eol,datasource_h2,download_folder,3)
