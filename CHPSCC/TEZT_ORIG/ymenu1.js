	var NoOffFirstLineMenus=10;			
	var LowBgColor='#06203B';			
	var LowSubBgColor='#06203B';			
	var HighBgColor='#515FA0';			
	var HighSubBgColor='#515FA0';			
	var FontLowColor='white';			
	var FontSubLowColor='white';			
	var FontHighColor='white';			
	var FontSubHighColor='white';			
	var BorderColor='white';			
	var BorderSubColor='white';			
	var BorderWidth=1;				
	var BorderBtwnElmnts=1;			
	var FontFamily="arial,comic sans ms,technical"	
	var FontSize=9;				
	var FontBold=1;				
	var FontItalic=0;				
	var MenuTextCentered='left';			
	var MenuCentered='left';			
	var MenuVerticalCentered='top';		
	var ChildOverlap=.2;				
	var ChildVerticalOverlap=.2;			
	var StartTop=155;				
	var StartLeft=10;				
	var VerCorrect=0;				
	var HorCorrect=0;				
	var LeftPaddng=3;				
	var TopPaddng=2;				
	var FirstLineHorizontal=0;			
	var MenuFramesVertical=1;			
	var DissapearDelay=1000;			
	var TakeOverBgColor=1;			
	var FirstLineFrame='navig';			
	var SecLineFrame='space';			
	var DocTargetFrame='space';			
	var TargetLoc='';				
	var HideTop=0;				
	var MenuWrap=1;				
	var RightToLeft=0;				
	var UnfoldsOnClick=0;			
	var WebMasterCheck=0;			
	var ShowArrow=1;				
	var KeepHilite=1;				
	var Arrws=['images/tri.gif',5,10,'images/tridown.gif',10,5,'images/trileft.gif',5,10];	
function BeforeStart(){return}
function AfterBuild(){return}
function BeforeFirstOpen(){return}
function AfterCloseAll(){return}
Menu1=new Array("Home","index.htm","",0,20,150);Menu2=new Array("About Us","aboutus.htm","",7);Menu2_1=new Array("History","partnershiphistory.htm","",0,20,150);Menu2_2=new Array("Membership","membership.htm","",0);Menu2_3=new Array("Staff","staff.htm","",0);Menu2_4=new Array("Newsletter","newsletter.htm","",0);Menu2_5=new Array("Contact Us","contactus.htm","",0);Menu2_6=new Array("Forums","http://www.chpscc.org/forums.htm","",0);Menu2_7=new Array("Wish List","wishlist.htm","",0);	
Menu3=new Array("Member Clinics","","",0);Menu4=new Array("Clinical Services","ClinicalServices.htm","",5);Menu4_1=new Array("Clinical Services","ClinicalServices.htm","",2,20,150);Menu4_1_1=new Array("Pharmacy Support","PharmacySupport.htm","",0,20,200);Menu4_1_2=new Array("Medical Interpretation Services","","",0);Menu4_2=new Array("Health Systems","","",2);Menu4_2_1=new Array("Health Systems Network","","",0,20,170);Menu4_2_2=new Array("HIPAA","","",0);Menu4_3=new Array("Committees","CScommittees.htm","",2);Menu4_3_1=new Array("Medical Directors Calendar","medicaldirectors.htm","",0,20,200);Menu4_3_2=new Array("Clinic Managers Calendar","clinicmanagers.htm","",0);
Menu4_4=new Array("Events","","",0);Menu4_5=new Array("Resources & Links","","",0);Menu5=new Array("Programs","http://www.chpscc.org/Programs.htm","",3);Menu5_1=new Array("Pediatric Weight Management","PWMP.htm","",0,20,190);Menu5_2=new Array("Community Diabetes Project","CommunityDiabetesProject.htm","",6);Menu5_2_1=new Array("Program Description","CDPProgramDescription.htm","",3,20,200);Menu5_2_1_1=new Array("History","CDPProgramDescription.htm","",0,20,150);
Menu5_2_1_2=new Array("What We Do","CDPProgramDescription.htm","",0);Menu5_2_1_3=new Array("Who We Are","","",0);
Menu5_2_2=new Array("Collaborations","DiabCollaborations.htm","",0);Menu5_2_3=new Array("Community Services","http://www.chpscc.org/CommunityServices.htm","",4);
Menu5_2_3_1=new Array("Diabetes & Chronic Conditions Self Management Classes","DCcSMC.htm","",0,20,340);
Menu5_2_3_2=new Array("Prevention Presentations","PreventionPresentations.htm","",0);Menu5_2_3_3=new Array("Supply Bank","SupplyBank.htm","",0);
Menu5_2_3_4=new Array("Support Group","SupportGroup.htm","",0);Menu5_2_4=new Array("Events","http://www.chpscc.org/DiabEvents.htm","",0);
Menu5_2_5=new Array("Resources & Links","http://www.chpscc.org/DiabResourcesandLinks.htm","",0);Menu5_2_6=new Array("How You Can Help","http://www.chpscc.org/Howcanyouhelp.htm","",0);
Menu5_3=new Array("Women's Health Partnership","http://www.chpscc.org/whpWomensHealthPartnership.htm","",6);Menu5_3_1=new Array("About Women's Health Partnership","http://www.chpscc.org/whpAboutWomensHealthpartnership.htm","",3,20,220);
Menu5_3_1_1=new Array("Program Description","http://www.chpscc.org/whpWomensHealthPartnership.htm","",0,20,160);Menu5_3_1_2=new Array("WHP Team Profiles","","",0);
Menu5_3_1_3=new Array("Advisory Committee List","","",0);Menu5_3_2=new Array("Health Education & Outreach","http://www.chpscc.org/whpHealthEducationandOutreach.htm","",3,20,210);
Menu5_3_2_1=new Array("Neighborhood Health Days","http://www.chpscc.org/whpNeighborhoodHealthDays.htm","",0,20,260);Menu5_3_2_2=new Array("Healthy Women, Health Choices Curriculum","http://www.chpscc.org/whpHWHCC.htm","",0);
Menu5_3_2_3=new Array("Women's Resource Directory","http://www.chpscc.org/whpWomensResourceDirectory.htm","",0);Menu5_3_3=new Array("Clinic Services & Resources","http://www.chpscc.org/whpClinicServicesandResources.htm","",4,20,210);
Menu5_3_3_1=new Array("Breast & Cervical Cancer Coordination and Navigation Program (CCAN)","http://www.chpscc.org/whpCCAN.htm","",0,20,420);Menu5_3_3_2=new Array("Breast & Cervical Cancer Treatment Program (BCCTP)","http://www.chpscc.org/whpBCCTP.htm","",0);
Menu5_3_3_3=new Array("Gabriella Patser Program","http://www.chpscc.org/whpGabriellaPatserProgram.htm","",0);Menu5_3_3_4=new Array("","","",0);
Menu5_3_4=new Array("Our Partnership","http://www.chpscc.org/whpOurPartnership.htm","",3,20,210);Menu5_3_4_1=new Array("Advisory Committee","","",0,20,150);
Menu5_3_4_2=new Array("Membership Profile","","",0);Menu5_3_4_3=new Array("General Membership","","",0);
Menu5_3_5=new Array("Upcoming Events","http://www.chpscc.org/whpEvents.htm","",0,20,210);Menu5_3_6=new Array("Funding & Training Opportunities","http://www.chpscc.org/whpFundingandtrainingOpportunities.htm","",0,20,210);
Menu6=new Array("Opportunities","opportunities.htm","",3);Menu6_1=new Array("Employment","employment.htm","",0,20,180);
Menu6_2=new Array("Internships","internships.htm","",0);Menu6_3=new Array("Volunteer Work","volunteer.htm","",0);Menu7=new Array("Training","","",5);Menu7_1=new Array("Diabetes","","",0,20,220);
Menu7_2=new Array("Health Education & Training Center","http://www.chpscc.org/HealthEducationandTrainingCenter.htm","",3);Menu7_2_1=new Array("HIV/Aids Education & Prevention","http://www.chpscc.org/HIVAidsEducationandPrevention.htm","",6,20,210);
Menu7_2_1_1=new Array("El Pueblo Against AIDS","http://www.chpscc.org/ElPuebloAgainstAIDS.htm","",0,20,260);Menu7_2_1_2=new Array("San Jose AIDS Education & Training Center","http://www.chpscc.org/SJAEtC.htm","",0);
Menu7_2_1_3=new Array("Resources (Local & National)","http://www.chpscc.org/Resourceslandn.htm","",0);Menu7_2_1_4=new Array("Testing","http://www.chpscc.org/HIVTesting.htm","",0);
Menu7_2_1_5=new Array("Fundraising Events","http://www.chpscc.org/FundraisingEvents.htm","",0);Menu7_2_1_6=new Array("Volunteers","http://www.chpscc.org/Volunteers.htm","",0);
Menu7_2_2=new Array("Emergency Preparedness","http://www.chpscc.org/EmergencyPreparedness.htm","",7);Menu7_2_2_1=new Array("Program Description","http://www.chpscc.org/calpenProgramDescription.htm","",0,20,210);
Menu7_2_2_2=new Array("Local Public Health Occurences","http://www.chpscc.org/LocalPublicHealthOccurences.htm","",0);Menu7_2_2_3=new Array("Trainings","http://www.chpscc.org/calpenTrainings.htm","",0);
Menu7_2_2_4=new Array("Faculty","http://www.chpscc.org/calpenFaculty.htm","",0);Menu7_2_2_5=new Array("Clinic Resources","http://www.chpscc.org/calpenClinicResources.htm","",0);
Menu7_2_2_6=new Array("Resources","http://www.chpscc.org/calpenResources.htm","",0);Menu7_2_2_7=new Array("Needs Assesment","http://www.chpscc.org/calpenNeedsAssesment.htm","",0);
Menu7_2_3=new Array("Health Professions Development","http://www.chpscc.org/HealthProfessionsDevelopment.htm","",4);Menu7_2_3_1=new Array("Local AHEC Program Overview","http://www.chpscc.org/LocalAHECProgramOverview.htm","",0,20,150);
Menu7_2_3_2=new Array("Healthy Futures","","",0);Menu7_2_3_3=new Array("Resources/Links","http://www.chpscc.org/HPResourcesandLinks.htm","",0);
Menu7_2_3_4=new Array("Trainings & Events","http://www.chpscc.org/HPTrainingandEvents.htm","",0);Menu7_3=new Array("Policy & Advocacy","","",0);
Menu7_4=new Array("Women's Health Partnership","http://www.chpscc.org/whpFundingandtrainingOpportunities.htm","",0);Menu7_5=new Array("Clinical Services","","",0);Menu8=new Array("Resources & Links","","",5);Menu8_1=new Array("Children's Health","","",0,20,220);
Menu8_2=new Array("Diabetes","http://www.chpscc.org/DiabResourcesandLinks.htm","",0);Menu8_3=new Array("Health Education & Training Center","","",0);
Menu8_4=new Array("Policy & Advocacy","","",0);Menu8_5=new Array("Women's Health Partnership","http://www.chpscc.org/whpClinicServicesandResources.htm","",0);
Menu9=new Array("Policy & Advocacy","","",6);Menu9_1=new Array("Grassroots Advocacy Activities","","",3,20,250);
Menu9_1_1=new Array("Patient Advocacy Program","","",0,20,180);Menu9_1_2=new Array("Voter registration Links","","",0);
Menu9_1_3=new Array("Find My Legislator Links","","",0);Menu9_2=new Array("Sta Clara County Legislative District Info","","",4);
Menu9_2_1=new Array("District Profiles","","",0,20,270);Menu9_2_2=new Array("Find My Legislator","","",0);
Menu9_2_3=new Array("Legislative District Chart of Member Clinics","","",0);Menu9_2_4=new Array("Legislative District Maps","","",0);
Menu9_3=new Array("News Flash","","",4);Menu9_3_1=new Array("Budget Information","","",0,20,230);
Menu9_3_2=new Array("Resources","","",0);Menu9_3_3=new Array("How to write a letter to your Legislator","","",0);
Menu9_3_4=new Array("Legislative District Maps","","",0);Menu9_4=new Array("Special Populations","","",3);Menu9_4_1=new Array("Women's Health","","",0,20,130);
Menu9_4_2=new Array("Children's Health","","",0);Menu9_4_3=new Array("Immigrant Health","","",0);
Menu9_5=new Array("Events","","",0);Menu9_6=new Array("Resources & Links","","",0);Menu10=new Array("Donate Now","http://www.chpscc.org/wishlist.htm","",0);	
