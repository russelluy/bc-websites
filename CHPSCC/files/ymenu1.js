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
Menu1=new Array("Home","index.htm","",0,20,150);Menu2=new Array("About Us","aboutus.htm","",7);Menu2_1=new Array("History","history1.htm","",0,20,150);Menu2_2=new Array("Membership","membership.htm","",0);Menu2_3=new Array("Staff","staff.htm","",0);Menu2_4=new Array("Newsletter","newsletter.htm","",0);Menu2_5=new Array("Contact Us","contactus.htm","",0);Menu2_6=new Array("Forums","Forums.htm","",0);Menu2_7=new Array("Wish List","wishlist.htm","",0);	
Menu3=new Array("Member Clinics","MemberClinics.htm","",0);Menu4=new Array("Clinical Services","ClinicalServices.htm","",5);Menu4_1=new Array("Clinical Services","ClinicalServices2.htm","",2,20,150);Menu4_1_1=new Array("Pharmacy Support","PharmacySupport.htm","",0,20,200);Menu4_1_2=new Array("Medical Interpretation Services","MedicalInterpretationServices.htm","",0);Menu4_2=new Array("Health Systems","HealthSystems.htm","",2);Menu4_2_1=new Array("Health Systems Network","HealthSystemsNetwork.htm","",0,20,170);Menu4_2_2=new Array("HIPAA","HIPAA.htm","",0);Menu4_3=new Array("Committees","Committees.htm","",2);Menu4_3_1=new Array("Medical Directors Calendar","MedicalDirectorsCalendar.htm","",0,20,200);Menu4_3_2=new Array("Clinic Managers Calendar","ClinicManagersCalendar.htm","",0);
Menu4_4=new Array("Events","Events.htm","",0);Menu4_5=new Array("Resources & Links","ResourcesandLinks.htm","",0);Menu5=new Array("Programs","Programs.htm","",3);Menu5_1=new Array("Children's Health Initiative","ChildrensHealthInitiative.htm","",0,20,190);Menu5_2=new Array("Community Diabetes Project","CommunityDiabetesProject.htm","",6);Menu5_2_1=new Array("Program Description","ProgramDescription.htm","",3,20,200);Menu5_2_1_1=new Array("History","History.htm","",0,20,150);
Menu5_2_1_2=new Array("What we do","Whatwedo.htm","",0);Menu5_2_1_3=new Array("Who we are","Whoweare.htm","",0);
Menu5_2_2=new Array("Collaborations","Collaborations.htm","",0);Menu5_2_3=new Array("Community Services","CommunityServices.htm","",4);
Menu5_2_3_1=new Array("Diabetes & Chronic conditions Self Management Classes","DCcSMC.htm","",0,20,340);
Menu5_2_3_2=new Array("Prevention Presentations","PreventionPresentations.htm","",0);Menu5_2_3_3=new Array("Supply Bank","SupplyBank.htm","",0);
Menu5_2_3_4=new Array("Support Group","SupportGroup.htm","",0);Menu5_2_4=new Array("Events","Events.htm","",0);
Menu5_2_5=new Array("Resources & Links","ResourcesandLinks.htm","",0);Menu5_2_6=new Array("How can you help","Howcanyouhelp.htm","",0);
Menu5_3=new Array("Women's Health Partnership","WomensHealthPartnership.htm","",6);Menu5_3_1=new Array("About Women's Health partnership","AboutWomensHealthpartnership.htm","",3,20,220);
Menu5_3_1_1=new Array("Program Description","ProgramDescription.htm","",0,20,160);Menu5_3_1_2=new Array("WHP Team Profiles","WHPTeamProfiles.htm","",0);
Menu5_3_1_3=new Array("Advisory Committee List","AdvisoryCommitteeList.htm","",0);Menu5_3_2=new Array("Health Education & Outreach","HealthEducationandOutreach.htm","",3,20,210);
Menu5_3_2_1=new Array("Nieghborhood Health Days","NieghborhoodHealthDays.htm","",0,20,260);Menu5_3_2_2=new Array("Healthy Women, Health Choices Curriculum","HWHCC.htm","",0);
Menu5_3_2_3=new Array("Women's Resource Directory","WomensResourceDirectory.htm","",0);Menu5_3_3=new Array("Clinic Services & Resources","ClinicServicesandResources.htm","",4,20,210);
Menu5_3_3_1=new Array("Success","Success.htm","",0,20,160);Menu5_3_3_2=new Array("CCAN","CCAN.htm","",0);
Menu5_3_3_3=new Array("Gabriella Patser Program","GabriellaPatserProgram.htm","",0);Menu5_3_3_4=new Array("BCCTP","BCCTP.htm","",0);
Menu5_3_4=new Array("Our Partnership","OurPartnership.htm","",3,20,210);Menu5_3_4_1=new Array("Advisory Committee","AdvisoryCommittee.htm","",0,20,150);
Menu5_3_4_2=new Array("Membership Profile","MembershipProfile.htm","",0);Menu5_3_4_3=new Array("General Membership","GeneralMembership.htm","",0);
Menu5_3_5=new Array("Events","Events.htm","",0,20,210);Menu5_3_6=new Array("Funding & training Opportunities","FundingandtrainingOpportunities.htm","",0,20,210);
Menu6=new Array("Opportunities","Opportunities.htm","",3);Menu6_1=new Array("Employment","Employment.htm","",0,20,180);
Menu6_2=new Array("Internships","Internships.htm","",0);Menu6_3=new Array("Volunteer Work","VolunteerWork.htm","",0);Menu7=new Array("Training & Events","TrainingandEvents.htm","",5);Menu7_1=new Array("Diabetes","Diabetes.htm","",0,20,220);
Menu7_2=new Array("Health Education & Training Center","HealthEducationandTrainingCenter.htm","",3);Menu7_2_1=new Array("HIV/Aids Education & Prevention","HIVAidsEducationandPrevention.htm","",6,20,210);
Menu7_2_1_1=new Array("El Pueblo Against AIDS","ElPuebloAgainstAIDS.htm","",0,20,260);Menu7_2_1_2=new Array("San Jose AIDS Education & training Center","SJAEtC.htm","",0);
Menu7_2_1_3=new Array("Resources (Local & National)","Resourceslandn.htm","",0);Menu7_2_1_4=new Array("Testing","Testing.htm","",0);
Menu7_2_1_5=new Array("Fundraising Events","FundraisingEvents.htm","",0);Menu7_2_1_6=new Array("Volunteers","Volunteers.htm","",0);
Menu7_2_2=new Array("Emergency Preparedness","EmergencyPreparedness.htm","",7);Menu7_2_2_1=new Array("Program Description","ProgramDescription.htm","",0,20,210);
Menu7_2_2_2=new Array("Local Public Health Occurences","LocalPublicHealthOccurences.htm","",0);Menu7_2_2_3=new Array("Trainings","Trainings.htm","",0);
Menu7_2_2_4=new Array("Faculty","Faculty.htm","",0);Menu7_2_2_5=new Array("Clinic Resources","ClinicResources.htm","",0);
Menu7_2_2_6=new Array("Resources","Resources.htm","",0);Menu7_2_2_7=new Array("Needs Assesment","NeedsAssesment.htm","",0);
Menu7_2_3=new Array("Health Professions Development","HealthProfessionsDevelopment.htm","",4);Menu7_2_3_1=new Array("Local AHEC Program Overview","LocalAHECProgramOverview.htm","",0,20,150);
Menu7_2_3_2=new Array("Healthy Futures","HealthyFutures.htm","",0);Menu7_2_3_3=new Array("Resources/Links","RL.htm","",0);
Menu7_2_3_4=new Array("Trainings & Events","TE.htm","",0);Menu7_3=new Array("Policy & Advocacy","PolicyandAdvocacy.htm","",0);
Menu7_4=new Array("Women's Health Partnership","WomensHealthPartnership.htm","",0);Menu7_5=new Array("Clinical Services","ClinicalServices.htm","",0);Menu8=new Array("Resources & Links","ResourcesandLinks.htm","",5);Menu8_1=new Array("Children's Health","ChildrensHealth.htm","",0,20,220);
Menu8_2=new Array("Diabetes","Diabetes.htm","",0);Menu8_3=new Array("Health Education & Training Center","HETC.htm","",0);
Menu8_4=new Array("Policy & Advocacy","PolicyandAdvocacy.htm","",0);Menu8_5=new Array("Women's Health Partnership","WomensHealthPartnership.htm","",0);
Menu9=new Array("Policy & Advocacy","PolicyandAdvocacy.htm","",6);Menu9_1=new Array("Grassroots Advocacy Activities","GrassrootsAdvocacyActivities.htm","",3,20,250);
Menu9_1_1=new Array("Patient Advocacy Program","PatientAdvocacyProgram.htm","",0,20,180);Menu9_1_2=new Array("Voter registration Links","VoterregistrationLinks.htm","",0);
Menu9_1_3=new Array("Find My Legislator Links","FindMyLegislatorLinks.htm","",0);Menu9_2=new Array("Sta Clara County Legislative District Info","SCCLDI.htm","",4);
Menu9_2_1=new Array("District Profiles","DistrictProfiles.htm","",0,20,270);Menu9_2_2=new Array("Find My Legislator","FindMyLegislator.htm","",0);
Menu9_2_3=new Array("Legislative District Chart of Member Clinics","LDCoMC.htm","",0);Menu9_2_4=new Array("Legislative District Maps","LegislativeDistrictMaps.htm","",0);
Menu9_3=new Array("News Flash","NewsFlash.htm","",4);Menu9_3_1=new Array("Budget Information","BudgetInformation.htm","",0,20,230);
Menu9_3_2=new Array("Resources","Resources.htm","",0);Menu9_3_3=new Array("How to write a letter to your Legislator","HowtowritealettertoyourLegislator.htm","",0);
Menu9_3_4=new Array("Legislative District Maps","LegislativeDistrictMaps.htm","",0);Menu9_4=new Array("Speical Populations","SpeicalPopulations.htm","",3);Menu9_4_1=new Array("Women's Health","WomensHealth.htm","",0,20,130);
Menu9_4_2=new Array("Children's Health","ChildrensHealth2.htm","",0);Menu9_4_3=new Array("Immigrant Health","ImmigrantHealth.htm","",0);
Menu9_5=new Array("Events","Events2.htm","",0);Menu9_6=new Array("Resources & Links","ResourcesandLinks2.htm","",0);Menu10=new Array("Donate Now","DonateNow.htm","",0);	
