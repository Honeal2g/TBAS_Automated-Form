function SendConfirmationEmail(e){
  var LastName = e.values[2];
  var FirstName = e.values[1];
  var UserEmail = e.values[3];
  var TestDate = e.values[6];
  var AFOQT = e.values[5];
  var TestTime;  
  var Subject = "TBAS Test Date";
  var BCC_Email = "uhawthorne01@manhattan.edu";
  var SignatureBlock = "\n\nVery Respectfully,\n\nULAN  O. HAWTHORNE, SSgt, USAF\n\nNCOIC, Administrative Management\n\nMANHATTAN COLLEGE\n\nLEO ENGINEERING BLDG"
  +"\n\n3825 Corlear Ave Rm 246\n\nRiverdale, NY 10463\n\nPhone: (718) 862-7202\n\nFax: (718) 862-7900\n\nSecondary Email: Ulan.Hawthorne@us.af.mil";
  
  for(var i=7; i!=37; i++){
    if(e.values[i] != ""){
      TestTime = e.values[i];      
      break;
    }
  }
   
  if(AFOQT == "Yes"){
    var Month = Convert_Appt(TestDate, 'M');
    var Day = Convert_Appt(TestDate, 'D');
    var S_Time = Convert_Time(TestTime, 'S');
    var E_Time = Convert_Time(S_Time, 'E');
      
    var Message = "Greetings, \n\n" + FirstName + " " + LastName + ",\n\nYou have successfully registered for your TBAS exam on: \n\n" + TestDate + " @ " + TestTime;
    var Directions = "\n\nTest Location: \n\nAir Force ROTC Detachment 560 is located at Manhattan College in Riverdale New York City. "
    + "Our office is in the Leo Engineering Building/Leo Hall, Room 246, 3825 Corlear Ave, Riverdale NY, 10463. "
    + "Please visit the following site for a campus map https://inside.manhattan.edu/_files/images/campus_map.pdf. If you are planning to drive on campus, you MUST park in P9: West 240th Street Parking. " 
    + "You will be issued a parking pass upon arrival. Inform the security guard that you are here for Air Force ROTC. "
    
    Message = Message + Directions +"\n\nIf you need to cancel or reschedule, please refer to my contact information below."+SignatureBlock;
    
    MailApp.sendEmail ({to: UserEmail, subject: Subject, body: Message}); // Can use htmlBody in future
    //createEvent(title, startTime, endTime, options)
    var event = CalendarApp.getDefaultCalendar().createEvent('TBAS Appointment for '+ FirstName +' '+ LastName, 
                                                             new Date(Month+' '+Day+', 2018 '+S_Time+' UTC'), 
                                                             new Date(Month+' '+Day+', 2018 '+E_Time+' UTC'),
                                                             {location: 'Air Force ROTC Detachment, 3825 Corlear Ave #246, Bronx, NY 10463, USA', 
                                                              guests: UserEmail, 
                                                              sendInvites: true});     
    event;    
  }
  
  if(AFOQT == "No"){
    var Message = "Greetings, \n\n" + FirstName + " " + LastName + ", The email is to notify you that your attempt to register for the TBAS exam via the Det560 online sign-up form was unsuccessful. " 
    + "In order to secure a test slot here at Detachment 560, you are required to have taken and received a passing AFOQT score. Once you have completed the pre-requisites listed on the website/form you may re-attempt the registration process.";
    
    Message = Message + SignatureBlock;
    MailApp.sendEmail ({to: UserEmail, subject: Subject, body: Message});     
  }
}

function Convert_Appt(Date, Flag){
  var Calender = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var NW = Number(Date.substring(0,2)); 
  var ND = Date.substring(String(Date).indexOf("/"), String(Date.length));
  
  if(Flag == 'M'){
    if(isNaN(NW)){
      var NW_FIX = Number(Date.substring(0,1));
      return Calender[NW_FIX-1];      
    }else{
      return Calender[NW-1];
    }
  }
  if(Flag == 'D'){
    if(String(ND).length == 7){
      return ND.substring(1,2);
    }else{
      return ND.substring(1,3);
    }      
  }
};

function Convert_Time(T, Flag){
  var UTC = ['13:00:00','15:00:00','17:00:00','19:00:00'];
  
  if(Flag == 'S'){
    if(String(T).length == "9:"){
      return UTC[0];
    }else{
      if(Number(T.substring(0,2)) == 11.0){
        return UTC[1];
      }else{
        return UTC[2];
      }
    }
  }
  if(Flag == 'E'){
    for(var i=0; i!=4; i++){
      if(String(T) == UTC[i]){
        return UTC[i+1];      
        i+=4;
      }
    }    
  }  
};

