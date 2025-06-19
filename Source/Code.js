/**
 * Career Starter Bootcamp - Access & Feedback Form Creator
 * Google Apps Script to create a form for undergraduate product testers
 */

function createCareerBootcampForm() {
  // Create a new form
  const form = FormApp.create('Career Starter Bootcamp – Access & Feedback Form');
  
  // Set form description
  form.setDescription('This form is designed for undergraduate product testers to provide access information and feedback for the Career Starter Bootcamp program.');
  
  // Set form settings
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(true);
  form.setShowLinkToRespondAgain(false);
  
  // Add Section 1: Participant Information
  const section1 = form.addSectionHeaderItem();
  section1.setTitle('Section 1: Participant Information');
  section1.setHelpText('Please provide your basic information for bootcamp access and communication.');
  
  // Full Name - Required text field
  const fullName = form.addTextItem();
  fullName.setTitle('Full Name');
  fullName.setHelpText('Enter your complete name as it appears on official documents');
  fullName.setRequired(true);
  
  // Email Address - Required email field
  const emailAddress = form.addTextItem();
  emailAddress.setTitle('Email Address');
  emailAddress.setHelpText('Enter a valid email address for bootcamp communications');
  emailAddress.setRequired(true);
  // Add email validation
  const emailValidation = FormApp.createTextValidation()
    .setHelpText('Please enter a valid email address')
    .requireTextIsEmail()
    .build();
  emailAddress.setValidation(emailValidation);
  
  // Phone Number - Required text field with validation
  const phoneNumber = form.addTextItem();
  phoneNumber.setTitle('Phone Number');
  phoneNumber.setHelpText('Enter your phone number (e.g., +234 xxx xxx xxxx)');
  phoneNumber.setRequired(true);
  // Add phone number validation (basic format)
  const phoneValidation = FormApp.createTextValidation()
    .setHelpText('Please enter a valid phone number')
    .requireTextMatchesPattern('^[+]?[0-9\\s\\-\\(\\)]{10,15}$')
    .build();
  phoneNumber.setValidation(phoneValidation);
  
  // University/Institution - Required text field
  const university = form.addTextItem();
  university.setTitle('University/Institution');
  university.setHelpText('Enter the name of your current university or educational institution');
  university.setRequired(true);
  
  // Course of Study - Required text field
  const courseOfStudy = form.addTextItem();
  courseOfStudy.setTitle('Course of Study');
  courseOfStudy.setHelpText('Enter your major/field of study (e.g., Computer Science, Business Administration)');
  courseOfStudy.setRequired(true);
  
  // Year of Study - Required dropdown
  const yearOfStudy = form.addListItem();
  yearOfStudy.setTitle('Year of Study');
  yearOfStudy.setHelpText('Select your current academic level');
  yearOfStudy.setRequired(true);
  yearOfStudy.setChoices([
    yearOfStudy.createChoice('100 Level'),
    yearOfStudy.createChoice('200 Level'),
    yearOfStudy.createChoice('300 Level'),
    yearOfStudy.createChoice('400 Level'),
    yearOfStudy.createChoice('500 Level'),
    yearOfStudy.createChoice('Others')
  ]);
  
  // Preferred Learning Mode - Required dropdown
  const learningMode = form.addListItem();
  learningMode.setTitle('Preferred Learning Mode');
  learningMode.setHelpText('Select your preferred mode of learning for the bootcamp');
  learningMode.setRequired(true);
  learningMode.setChoices([
    learningMode.createChoice('Online'),
    learningMode.createChoice('In-Person'),
    learningMode.createChoice('Hybrid')
  ]);
  
  // Add Additional Questions Section for Product Testing Feedback
  const section2 = form.addSectionHeaderItem();
  section2.setTitle('Section 2: Product Testing & Feedback');
  section2.setHelpText('As a product tester, your feedback is valuable for improving our bootcamp experience.');
  
  // Previous Experience
  const previousExperience = form.addMultipleChoiceItem();
  previousExperience.setTitle('Have you participated in similar career development programs before?');
  previousExperience.setRequired(true);
  previousExperience.setChoices([
    previousExperience.createChoice('Yes'),
    previousExperience.createChoice('No')
  ]);
  
  // Areas of Interest
  const areasOfInterest = form.addCheckboxItem();
  areasOfInterest.setTitle('Which career areas are you most interested in? (Select all that apply)');
  areasOfInterest.setRequired(true);
  areasOfInterest.setChoices([
    areasOfInterest.createChoice('Technology & Software Development'),
    areasOfInterest.createChoice('Business & Entrepreneurship'),
    areasOfInterest.createChoice('Marketing & Digital Marketing'),
    areasOfInterest.createChoice('Finance & Accounting'),
    areasOfInterest.createChoice('Data Science & Analytics'),
    areasOfInterest.createChoice('Design & Creative Arts'),
    areasOfInterest.createChoice('Project Management'),
    areasOfInterest.createChoice('Others')
  ]);
  
  // Expectations
  const expectations = form.addParagraphTextItem();
  expectations.setTitle('What are your main expectations from the Career Starter Bootcamp?');
  expectations.setHelpText('Please describe what you hope to achieve or learn from this program');
  expectations.setRequired(true);
  
  // Availability
  const availability = form.addMultipleChoiceItem();
  availability.setTitle('How many hours per week can you dedicate to the bootcamp?');
  availability.setRequired(true);
  availability.setChoices([
    availability.createChoice('Less than 5 hours'),
    availability.createChoice('5-10 hours'),
    availability.createChoice('10-15 hours'),
    availability.createChoice('15-20 hours'),
    availability.createChoice('More than 20 hours')
  ]);
  
  // Additional Comments
  const additionalComments = form.addParagraphTextItem();
  additionalComments.setTitle('Additional Comments or Suggestions');
  additionalComments.setHelpText('Any other information you would like to share or suggestions for the bootcamp program');
  additionalComments.setRequired(false);
  
  // Create a new spreadsheet for responses
  const spreadsheet = SpreadsheetApp.create('Career Bootcamp Form Responses');
  const spreadsheetId = spreadsheet.getId();
  
  // Set up response collection to the new spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  
  // Get the form URL and ID
  const formUrl = form.getPublishedUrl();
  const formId = form.getId();
  
  // Log the details
  console.log('Form created successfully!');
  console.log('Form ID: ' + formId);
  console.log('Form URL: ' + formUrl);
  console.log('Edit URL: ' + form.getEditUrl());
  console.log('Spreadsheet ID: ' + spreadsheetId);
  console.log('Spreadsheet URL: ' + spreadsheet.getUrl());
  
  // Return form details
  return {
    id: formId,
    url: formUrl,
    editUrl: form.getEditUrl(),
    title: form.getTitle(),
    spreadsheetId: spreadsheetId,
    spreadsheetUrl: spreadsheet.getUrl()
  };
}

/**
 * Function to set up form notifications and auto-responses
 */
function setupFormNotifications() {
  // This function can be called after creating the form to set up additional features
  const forms = FormApp.getActiveForm();
  
  if (forms) {
    // Set confirmation message
    forms.setConfirmationMessage(
      'Thank you for registering for the Career Starter Bootcamp! ' +
      'We have received your information and will contact you soon with further details. ' +
      'Keep an eye on your email for updates and bootcamp materials.'
    );
    
    console.log('Form notifications configured successfully!');
  } else {
    console.log('No active form found. Please run createCareerBootcampForm() first.');
  }
}

/**
 * Function to get form responses and process them
 */
function processFormResponses() {
  try {
    const form = FormApp.getActiveForm();
    if (!form) {
      console.log('No active form found.');
      return;
    }
    
    const responses = form.getResponses();
    console.log(`Total responses received: ${responses.length}`);
    
    // Process each response
    responses.forEach((response, index) => {
      const itemResponses = response.getItemResponses();
      console.log(`\nResponse ${index + 1}:`);
      console.log(`Submitted at: ${response.getTimestamp()}`);
      console.log(`Email: ${response.getRespondentEmail()}`);
      
      itemResponses.forEach(itemResponse => {
        console.log(`${itemResponse.getItem().getTitle()}: ${itemResponse.getResponse()}`);
      });
    });
    
  } catch (error) {
    console.error('Error processing responses:', error);
  }
}

/**
 * Function to run everything at once
 */
function main() {
  const formDetails = createCareerBootcampForm();
  console.log('\n=== FORM CREATED SUCCESSFULLY ===');
  console.log('Please save these details:');
  console.log('Form ID:', formDetails.id);
  console.log('Public URL:', formDetails.url);
  console.log('Edit URL:', formDetails.editUrl);
  console.log('Spreadsheet ID:', formDetails.spreadsheetId);
  console.log('Spreadsheet URL:', formDetails.spreadsheetUrl);
  console.log('\nNext steps:');
  console.log('1. Open the Edit URL to customize the form further if needed');
  console.log('2. Share the Public URL with your undergraduate product testers');
  console.log('3. Run setupFormNotifications() to configure confirmation messages');
  console.log('4. Run processFormResponses() to view submitted responses');
}