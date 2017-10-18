// Set Global variables for quick addressing

var LOCAL = SpreadsheetApp.getActiveSpreadsheet();
var SHEET = LOCAL.getActiveSheet();
var RANGE = LOCAL.getActiveRange();
var MAIN = LOCAL.getSheetByName('MAIN');
var DATA = LOCAL.getSheetByName('Data');
var LUKE = "luke@lnmsupplies.com";
var MICH = "michelle@lnmsupplies.com";
  
  function onOpen(){  
   var user = Session.getActiveUser().getEmail();
      Logger.log(user);   
   var range = MAIN.getRange(1, 1);
      MAIN.setActiveRange(range);

   
       if(user == LUKE){
          var office = 'Send to New Customer';
          var New = 'openNew';
          var reset = 'Reset Sheet';
          var resetFunc = 'resetSheet';
          var admin = "Luke";
             createMenu();
             return;
       }
       else{
          var office = '';
          var New = '';
          var reset = '';
          var resetFunc = '';
          createMenu();
             return;
             }
         function createMenu(){     
         var ui = SpreadsheetApp.getUi();
            ui.createMenu('L&M Supplies')
            .addItem('Show all Product Lines','showProducts')
            .addItem('Hide all Product Lines','hideProducts')
            .addSeparator()
            .addItem('Change Contact Details','openForm')
            .addItem('Review Account Details','showInfo')
            .addSeparator()
            .addItem('Submit','onButtonClick')
            .addItem(office,New)
            .addItem(reset,resetFunc)
            .addToUi()
              openLogin();
              showDates();
              hideProducts();
        };
    };
    
  function showDates(){
   var oldDateRange = DATA.getRange(10,4);
    var oldDate = oldDateRange.getValue();
    var lastLogin = MAIN.getRange(10,8).setValue(oldDate);
    var currentDateRange = DATA.getRange(9,8);
    var currentDate = currentDateRange.getValue();
    oldDateRange.setValue(currentDate);
  };
  
  function showProducts(){
   var range = MAIN.getRange(4,1);
    DATA.getRange(2,2).setValue('Show');
    showPoorBoy();
    DATA.getRange(8,2).setValue('Show');
    showRootbeer();
    DATA.getRange(9,2).setValue('Show');
    show40Degrees();   
    DATA.getRange(3,2).setValue('Show');
    showVapeDroid();
    DATA.getRange(4,2).setValue('Show');
    showElfin();
    DATA.getRange(5,2).setValue('Show');
    showAttys();
    DATA.getRange(6,2).setValue('Show');
    showMal();
    DATA.getRange(7,2).setValue('Show');
    showSlass();
    DATA.getRange(10,2).setValue('Show');
    showLancer();
    DATA.getRange(11,2).setValue('Show');
    showGrip();
    DATA.getRange(12,2).setValue('Show');
    showPyxy(); 
    DATA.getRange(13,2).setValue('Show');
    showBeach(); 
    DATA.getRange(14,2).setValue('Show');
    showAngel();
   
   
    MAIN.setActiveRange(range);
     return

  };
  
  function hideProducts(){
    DATA.getRange(1,2).setValue('Hide');
    showInfo();
    DATA.getRange(2,2).setValue('Hide');
    showPoorBoy();
    DATA.getRange(3,2).setValue('Hide');
    showVapeDroid();
    DATA.getRange(4,2).setValue('Hide');
    showElfin();
    DATA.getRange(5,2).setValue('Hide');
    showAttys();
    DATA.getRange(6,2).setValue('Hide');
    showMal();   
    DATA.getRange(7,2).setValue('Hide');
    showSlass();
    DATA.getRange(8,2).setValue('Hide');
    showRootbeer();
    DATA.getRange(9,2).setValue('Hide');
    show40Degrees();
    DATA.getRange(10,2).setValue('Hide');
    showLancer();
    DATA.getRange(11,2).setValue('Hide');
    showGrip();
    DATA.getRange(12,2).setValue('Hide');
    showPyxy(); 
    DATA.getRange(13,2).setValue('Hide');
    showBeach();
    DATA.getRange(14,2).setValue('Hide');
    showAngel();
   
  };
  
  function showhideButton() {
  var showCount = DATA.getRange(1,3).getValue();
  var hideCount = DATA.getRange(1,4).getValue();
  if(hideCount >= showCount) {
  hideProducts();
  }else{
  showProducts();
  } };
  
  function sendForm() {
    Logger.log('sendForm ran!');
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showModalDialog(html,"Thank You For your Order");
  };
  
  /////////////////////////////////////////////////////
  /// Check if the customer is already set up  
  ///////////////////////////////////////////////////
  
  function openLogin() {
    var email = MAIN.getRange('emailRange').getValue();
      Logger.log(email);
      if( email == "your@daily_e-mail.here"){
        Logger.log("open Login Called!");
        var html = HtmlService.createHtmlOutputFromFile('openDialog').setWidth(300).setHeight(600);
        SpreadsheetApp.getUi().showModalDialog(html,"Please enter your email address");
    };
  };
    
function processUser(userObject){
    var uEmail = userObject.userEmail;
    Logger.log("processUser has been called");
    var s = SpreadsheetApp.openById("1rkeoMHv8TErN5y7Jnh8UBvZ7H8fKjA61yPrjcbLvOE0").getSheetByName("CustomerIndex");
    var dataRange = s.getRange(2,2,s.getLastRow(),s.getLastColumn());
    var data = dataRange.getDisplayValues();
  for(i=0; i < data.length; i++){
    Logger.log("User Email"+ uEmail);  
    var row = data[i];
    var company = row[0];
    var first = row[1];
    var last = row[2];
    var email = row[3];
    var phone = row[4];
    var addy1 = row[6];
    var addy2 = row[7];
    var city = row[8];
    var prov = row[9];
    var postal = row[10];
    var country = row[11]; 
    if (uEmail === email){  
      if(company !=""){MAIN.getRange('CompName').setValue(company)};  
      if(first !=""){MAIN.getRange('firstName').setValue(first)};
      if(last !=""){MAIN.getRange('lastName').setValue(last)};
      if(email !=""){MAIN.getRange('emailRange').setValue(email)};
      if(phone !=""){MAIN.getRange('phoneRange').setValue(phone)};
      if(addy1 !=""){MAIN.getRange('addyOne').setValue(addy1)};
      if(addy2 !=""){MAIN.getRange('addyTwo').setValue(addy2)};
      if(city !=""){MAIN.getRange('cityRange').setValue(city)};
      if(prov !=""){MAIN.getRange('provRange').setValue(prov)};
      if(postal !=""){MAIN.getRange('postRange').setValue(postal)};
      if(country !=""){MAIN.getRange('countryRange').setValue(country)}; 
    };
  };
   if (uEmail != MAIN.getRange('emailRange').getValue()){
   openForm();};

};

///////////////////////////////////////////////////
/// Checking that the contact info is filled out 
///////////////////////////////////////////////////
function onEdit(){
                  if( MAIN.getRange('CompName') == ""){setValue("Company Name")};
                  if( MAIN.getRange('firstName') == ""){setValue("Enter First Name")};
                  if( MAIN.getRange('lastName') == ""){setValue("Enter Last Name")};
                  if( MAIN.getRange('emailRange') == ""){setValue("your@daily_e-mail.here")};
                  if( MAIN.getRange('phoneRange') == ""){setValue("(123)-456-7890")};
                  if( MAIN.getRange('addyOne') == ""){setValue("Shipping St #")};
                  if( MAIN.getRange('addyTwo') == ""){setValue("Shipping unit #")};
                  if( MAIN.getRange('cityRange') == ""){setValue("Shipping City")};
                  if( MAIN.getRange('provRange') == ""){setValue("Province/State")};
                  if( MAIN.getRange('postRange') == ""){setValue("Postal/Zip COde")};
                  if( MAIN.getRange('countryRange') == ""){setValue("Company Name")};
                   };
                   
function resetSheet(){
                  MAIN.getRange('Contact').clearContent();
                  MAIN.getRange('ShippingInfo').clearContent();
                  MAIN.getRange('CompName').setValue("Company Name");
                  MAIN.getRange('firstName').setValue("Enter First Name");;
                  MAIN.getRange('lastName').setValue("Enter Last Name");
                  MAIN.getRange('emailRange').setValue("your@daily_e-mail.here");
                  MAIN.getRange('phoneRange').setValue("(123)-456-7890");
                  MAIN.getRange('addyOne').setValue("Shipping St #");
                  MAIN.getRange('addyTwo').setValue("Shipping unit #");
                  MAIN.getRange('cityRange').setValue("In'ON'2 letter standard");
                  MAIN.getRange('provRange').setValue("Province/State");
                  MAIN.getRange('postRange').setValue("Postal/Zip Code");
                  MAIN.getRange('countryRange').setValue("Country Name"); 
                  };


/////////////////////////////////////////////////
// For returning the Contact info to the form 
//////////////////////////////////////////////////
function openForm() {
    Logger.log('openForm ran!');
    var html = HtmlService.createHtmlOutputFromFile('index').setWidth(900).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html,"Please fill in the fields below. If your Updating then only fill the field need and leave the other blank");
    return;
    };
    
function processForm(formObject) {
  var formCompName = formObject.argCompName;
  var formFirstN = formObject.argFirstName;
  var formLastN= formObject.argLastName;
  var formEmail = formObject.argEmail;
  var formPhone = formObject.argPhone;
  var formAddress1 = formObject.argAddress1;
  var formAddress2 = formObject.argAddress2;
  var formCity = formObject.argCity;
  var formProvince = formObject.argProvince;
  var formCode = formObject.argPostalCode;
  var formCountry = formObject.argCountry;
    if( formCompName != ""){MAIN.getRange('CompName').setValue(formCompName)};
    if( formFirstN != ""){MAIN.getRange('firstName').setValue(formFirstN)};
    if( formLastN != ""){MAIN.getRange('lastName').setValue(formLastN)};
    if( formEmail != ""){MAIN.getRange('emailRange').setValue(formEmail)};
    if( formPhone != ""){MAIN.getRange('phoneRange').setValue(formPhone)};
    if( formAddress1 != ""){MAIN.getRange('addyOne').setValue(formAddress1)};
    if( formAddress2 != ""){MAIN.getRange('addyTwo').setValue(formAddress2)};
    if( formCity != ""){MAIN.getRange('cityRange').setValue(formCity)};
    if( formProvince != ""){MAIN.getRange('provRange').setValue(formProvince)};
    if( formCode != ""){MAIN.getRange('postRange').setValue(formCode)};
    if( formCode != ""){MAIN.getRange('countryRange').setValue(formCountry)};
   Logger.log("formCompName: " + formCompName);
 };
 

///////////////////////////////////////////
/// For Sending to a new customer
///////////////////////////////////////////
  
  function openNew(){
    var email = MAIN.getRange('emailRange').getValue();
     Logger.log(email);
    if( email == "your@daily_e-mail.here"){
      Logger.log("open Login Called!");
      var html = HtmlService.createHtmlOutputFromFile('newCustomer').setWidth(300).setHeight(600);
      SpreadsheetApp.getUi().showModalDialog(html,"Please enter your email address");
    };
  };
  
  function createNew(newObject){
    var firstName = newObject.userFirstName;
    var uEmail = newObject.userEmail;
    Logger.log("processUser has been called");
    var s = SpreadsheetApp.openById("1rkeoMHv8TErN5y7Jnh8UBvZ7H8fKjA61yPrjcbLvOE0").getSheetByName("CustomerIndex");
    var dataRange = s.getRange(2,2,s.getLastRow(),s.getLastColumn());
    var data = dataRange.getDisplayValues();
    for(i=0; i < data.length; i++){
      var row = data[i];
      var company = row[0];
      var first = row[1];
      var last = row[2];
      var email = row[3];
      var phone = row[4];
      var addy1 = row[6];
      var addy2 = row[7];
      var city = row[8];
      var prov = row[9];
      var postal = row[10];
      var country = row[11]; 
      if (uEmail === email){ 
          if(company !=""){MAIN.getRange('CompName').setValue(company)};  
          if(first !=""){MAIN.getRange('firstName').setValue(first)};
          if(last !=""){MAIN.getRange('lastName').setValue(last)};
          if(email !=""){MAIN.getRange('emailRange').setValue(email)};
          if(phone !=""){MAIN.getRange('phoneRange').setValue(phone)};
          if(addy1 !=""){MAIN.getRange('addyOne').setValue(addy1)};
          if(addy2 !=""){MAIN.getRange('addyTwo').setValue(addy2)};
          if(city !=""){MAIN.getRange('cityRange').setValue(city)};
          if(prov !=""){MAIN.getRange('provRange').setValue(prov)};
          if(postal !=""){MAIN.getRange('postRange').setValue(postal)};
          if(country !=""){MAIN.getRange('countryRange').setValue(country)}; 
      };
    };
        var ssUrl = LOCAL.getId();
        var aManager = LUKE;
        var date = MAIN.getRange(9,7).getValue();
        var contact = MAIN.getRange('firstName').getValue();
        var email =  MAIN.getRange('emailRange').getValue();
        var companyName = MAIN.getRange('CompName').getValue();
        var curName = companyName; 
        var dName = LOCAL.getId();
        DriveApp.getFoldersByName("04-Customers").next()
        try {DriveApp.getFoldersByName(companyName).next()}
          catch(e)
            { var compName = MAIN.getRange('CompName').getValue();
              DriveApp.createFolder(compName) 
              };
        var MyVendors = DriveApp.getFoldersByName(companyName).next();
        var cusfile = curName +"-"+companyName+"-Master Copy";
        var file = DriveApp.getFileById(dName).makeCopy(cusfile,MyVendors).addEditors([email, LUKE]);
        var ssNew = file.getDownloadUrl();
        //var remove = file.removeEditor(EmailRange,aManager);
        var reset1 = LOCAL.getRangeByName("Contact").clearContent();
        var reset2 = LOCAL.getRangeByName("ShippingInfo").clearContent();
        var reset3 = LOCAL.rename(curName+"-Master Copy");
  };

//////////////////////////////////////////////
// Submitting the form 
/////////////////////////////////////////////
  function onButtonClick(){
    var totalData = LOCAL.getSheetByName("Order").getDataRange();
    var rows = totalData.getLastRow()
    var luke = "Luke@LnMSupplies.com";
    var date = MAIN.getRange(9,8).getValue();
    var companyNameRange = MAIN.getRange('CompName');
    var contact = MAIN.getRange('firstName').getValue();
    var EmailRange = MAIN.getRange('emailRange').getValue();
    // var PhoneRange = s.getRange(6,7);
    // var AddressRange = s.getRange(8,4);
    var curName = MAIN.getRange(1,10).getValue();     
    var CompanyName = companyNameRange.getValue();
    ss.rename(curName +"New Order From"+CompanyName+" on  "+date);
    var fileName = LOCAL.getName();
    Logger.log(fileName);
    var dName = LOCAL.getId() ;
    var file = DriveApp.getFileById(dName).makeCopy(fileName).addEditors([luke, EmailRange]);
    if (totalData.length >= 12) {
      // Add the ".csv" extension to the file name
      var  csvfileName = fileName + ".csv";
      // Convert the range data to CSV format
      var csvFile = convertRangeToCsvFile_(fileName);    
      var csvLoc =  DriveApp.createFile(fileName, csvFile);
      //var csvBlob = UrlFetchApp.fetch(csvLoc).getBlob();
      // var zipBlob = Utilities.zip([csvBlob],fileName + ".zip");
      // DriveApp.createFile(zipBlob);
      Logger.log("Storage Space used: " + DriveApp.getStorageUsed());
      var vEmail = Session.getActiveUser().getEmail();
      Logger.log(vEmail);
      var subject = "New Order From:"+CompanyName+date;
      GmailApp.sendEmail([LUKE, EmailRange],subject,"Please accept this .CSV submission and the Spreadsheet shared with you as my lastest Order and invoice me accordingly. Sincerely:"+Contact,  {
                         attachments: [csvLoc]
                         });
     
      ss.getRangeByName("BeaverTails").clearContent();
      ss.getRangeByName("CrumbMuffin").clearContent();
      ss.getRangeByName("KissTheSky").clearContent();
      ss.getRangeByName("MadMonkey").clearContent();
      ss.getRangeByName("TropicsOasis").clearContent();
      ss.getRangeByName("WildTropics").clearContent();
      ss.getRangeByName("Hurricane").clearContent();
      ss.getRangeByName("VapeDroidOdr").clearContent();
      ss.getRangeByName("ElfinOdr").clearContent();
      ss.getRangeByName("AttyOdr").clearContent();
      ss.getRangeByName("MalOdr").clearContent();
      ss.getRangeByName("SlassOdr").clearContent();
      ss.getRangeByName("Contact").clearContent();
      ss.getRangeByName("ShippingInfo").clearContent();
      ss.getRangeByName("gripOdr").clearContent();
      ss.getRangeByName("lancerOdr").clearContent();
      
      ss.rename(curName+"-Master Copy");
      Browser.msgBox("Thank you for your order! A copy of the order has been saved in Google drive of "
                     + vEmail+ " which can be found at www.drive.google.com . "
                     +EmailRange+". The Master form has been cleared for your next order!"+
                     "You should receive an email shortly with your invoice")
    }
    else {
      Browser.msgBox("Error: Please enter a CSV file name.");
    };
    SpreadsheetApp.flush();
  };

/**
 * For generating the csv file
 */
 /*
 function convertRangeToCsvFile_(csvFileName) {
    // Get the selected range in the spreadsheet
    var ws = SpreadsheetApp.getActiveSpreadsheet().getRange("TotalsData") ;
    try {
      var data = ws.getValues();
      var csvFile = undefined;
      */
/** 
 * Loop through the data in the range and build a string with the CSV data 
 * 
 */
 /*
 if (data.length > 1) {
 var csv = "";
        for (var row = 0; row < data.length; row++) {
          for (var col = 0; col < data[row].length; col++) {
            if (data[row][col].toString().indexOf(",") != -1) {
              data[row][col] = "\"" + data[row][col] + "\"";
            }
          }
          
          // Join each row's columns
          // Add a carriage return to end of each row, except for the last one
          if (row < data.length-1) {
            csv += data[row].join(",") + "\r\n";
          }
          else {
            csv += data[row];
          }
        }
        csvFile = csv;
      }
      return csvFile;
    }
    catch(err) {
        Logger.log(err);
      Browser.msgBox(err);
      
    }  
  }
  */
  
  
 ////////////////////////////////////////////////////// 
 // SHOW HIDDE FUNCTIONS
 //////////////////////////////////////////////////////
 /**
  * Hide Or Collapse the select your lines rows
  */

function showInfo(){ 
 var range = MAIN.getRange(4, 4);
  var selran = DATA.getRange(1,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(3,10);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(3,10);
       selran.setValue('Hide');
         MAIN.setActiveRange(range);
     return
     };
   };

function showPoorBoy(){ 
  var poorNum = MAIN.getRange('poorNum');
  var poorRow = poorNum.getRow();
  var numRows = poorNum.getValue()+5
  var selran = DATA.getRange(2,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(poorRow-2,numRows);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(poorRow-2,numRows);
       selran.setValue('Hide');
     return
     };
};

function showRootbeer(){ 
  var rbNum = MAIN.getRange('rbNum');
  var rbRow = rbNum.getRow();
  var numRows = rbNum.getValue()+4
  var selran = DATA.getRange(8,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(rbRow-2,numRows);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(rbRow-2,numRows);
       selran.setValue('Hide');
     return
     };
};

function show40Degrees(){  
  var fortyNum = MAIN.getRange('fortyNum');
  var fortyRow = fortyNum.getRow();
  var numRows = fortyNum.getValue()+4
  var selran = DATA.getRange(9,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(fortyRow-2,numRows);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(fortyRow-2,numRows);
       selran.setValue('Hide');
     return
     };
};

function showPyxy(){  
  var pyxyNum = MAIN.getRange('pyxyNum');
  var pyxyRow = pyxyNum.getRow();
  var numRows = pyxyNum.getValue()+4
  var selran = DATA.getRange(12,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(pyxyRow-2,numRows);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(pyxyRow-2,numRows);
       selran.setValue('Hide');
     return
     };
  };

function showBeach(){  
  var beachNum = MAIN.getRange('beachNum');
  var beachRow = beachNum.getRow();
  var numRows = beachNum.getValue()+4
  var selran = DATA.getRange(13,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(beachRow-2,numRows);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(beachRow-2,numRows);
       selran.setValue('Hide');
     return
     };
  };

function showAngel(){  
  var angelNum = MAIN.getRange('angelNum');
  var angelRow = angelNum.getRow();
  var numRows = angelNum.getValue()+4
  var selran = DATA.getRange(14,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(angelRow-2,numRows);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(angelRow-2,numRows);
       selran.setValue('Hide');
     return
     };
  };

function sbodyHide(){
  showElfin();
  showAttys();
  showMal();
  showSlass();
};

function showLancer(){ 
  var lancerNum = MAIN.getRange('lancerNum');
  var lancerRow = lancerNum.getRow();
  var numRows = lancerNum.getValue()+4
  Logger.log(numRows);
  var selran = DATA.getRange(10,2);
  var select = selran.getValue();
   if(select == 'Hide'){
     MAIN.hideRows(lancerRow-2,13);
     selran.setValue('Show');
   return
    }; 
   if(select == 'Show'){
     MAIN.showRows(lancerRow-2,numRows);
     selran.setValue('Hide');
   return
     };
   }; 
     
function showGrip(){ 
  var gripNum = MAIN.getRange('gripNum');
  var gripRow = gripNum.getRow();
  var numRows = gripNum.getValue()+4
  Logger.log(numRows);
  var selran = DATA.getRange(11,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(gripRow-2,13);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(gripRow-2,numRows);
       selran.setValue('Hide');
     return
       };
     }; 

function showVapeDroid(){ 
  var droidNum = MAIN.getRange('droidNum');
  var droidRow = droidNum.getRow();
  var numRows = droidNum.getValue()+4
  Logger.log(numRows);
  var selran = DATA.getRange(3,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(droidRow-2,13);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(droidRow-2,numRows);
       selran.setValue('Hide');
     return
       };
     }; 


function showElfin(){ 
  var elfNum = MAIN.getRange('elfNum');
  var elfRow = elfNum.getRow();
  var numRows = elfNum.getValue()+4
  var selran = DATA.getRange(4,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(elfRow-2,17);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(elfRow-2,numRows);
       selran.setValue('Hide');
     return
       };
     };
    
function showAttys(){ 
  var attyNum = MAIN.getRange('attyNum');
  var attyRow = attyNum.getRow();
  var numRows = attyNum.getValue()+4;
  var selran = DATA.getRange(5,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(attyRow-2,20);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(attyRow-2,numRows);
       selran.setValue('Hide');
     return
       };
     };
     
     function showMal(){ 
  var malNum = MAIN.getRange('malNum');
  var malRow = malNum.getRow();
  var numRows = malNum.getValue()+4;
  var selran = DATA.getRange(6,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(malRow-2,14);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(malRow-2,numRows);
       selran.setValue('Hide');
     return
       };
     };
     function showSlass(){ 
  var slass = MAIN.getRange('sLassNum');
  var sRow = slass.getRow();
  var numRows = slass.getValue()+4;
  var selran = DATA.getRange(7,2);
  var select = selran.getValue();
     if(select == 'Hide'){
       MAIN.hideRows(sRow-2,13);
       selran.setValue('Show');
     return
     }; 
     if(select == 'Show'){
       MAIN.showRows(sRow-2,numRows);
       selran.setValue('Hide');
     return
     };
     };
     
/*function showSidebar() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('Page.html')
  html.setSandboxMode(HtmlService.SandboxMode.IFRAME); 
  html.setTitle("ZeeroSmoke Ordering"); 
  ui.showSidebar(html);
};
*/
///////////////////////////////////////////////////////////////////////////////////
////// Test email function .
/////////////////////////////////////////////////////////////////////////////////
/*
function testSchemas(){
  var htmlBody = HtmlService.createHtmlOutputFromFile('mail_template').getContent();

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'Test Email markup - ' + new Date(),
    htmlBody: htmlBody,
  });
};
*/
///////////////////////////////////////////////////////////////////////////////////
////// Test email function .
/////////////////////////////////////////////////////////////////////////////////


