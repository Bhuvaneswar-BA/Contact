import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { EmailClient } from "@azure/communication-email";

// Move these to environment variables in Azure Function App settings
const CONNECTION_STRING = process.env.ACS_CONNECTION_STRING as string;
const SENDER_ADDRESS = process.env.SENDER_EMAIL_ADDRESS || "DoNotReply@2dde48cf-f3cb-436b-838a-1c27aa0e1c0c.azurecomm.net";
const RECIPIENT_ADDRESSES = ["brad@bullattorneys.com","web@bullattorneys.com"];

interface ContactFormData {
  firstName: string;
  lastName: string;
  email: string;
  phone: string;
  zipCode: string;
  caseType: string;
  description?: string;
}

export async function contactForm(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  context.log('ContactForm function processing started');
  
  try {
    context.log('Parsing request body');
    const requestBody = await request.text();
    context.log(`Request body: ${requestBody}`);
    
    let formData: ContactFormData;
    try {
      formData = JSON.parse(requestBody) as ContactFormData;
      context.log('Successfully parsed JSON data', formData);
    } catch (parseError) {
      context.error('Failed to parse request body as JSON:', parseError);
      return {
        status: 400,
        jsonBody: { error: "Invalid JSON in request body" }
      };
    }

    // Validate required fields
    context.log('Validating required fields');
    if (!formData.firstName || !formData.lastName || !formData.email || !formData.phone || !formData.zipCode || !formData.caseType) {
      const missingFields = [];
      if (!formData.firstName) missingFields.push('firstName');
      if (!formData.lastName) missingFields.push('lastName');
      if (!formData.email) missingFields.push('email');
      if (!formData.phone) missingFields.push('phone');
      if (!formData.zipCode) missingFields.push('zipCode');
      if (!formData.caseType) missingFields.push('caseType');
      
      context.log(`Validation failed. Missing fields: ${missingFields.join(', ')}`);
      return {
        status: 400,
        jsonBody: { error: "All required fields must be filled out", missingFields }
      };
    }

    // Initialize Email Client
    context.log('Initializing Email Client');
    context.log(`Using connection string starting with: ${CONNECTION_STRING.substring(0, 30)}...`);
    context.log(`Sender address: ${SENDER_ADDRESS}`);
    context.log(`Recipient addresses: ${RECIPIENT_ADDRESSES.join(', ')}`);
    
    try {
      const emailClient = new EmailClient(CONNECTION_STRING);
      context.log('Email client initialized successfully');

      const { firstName, lastName, email, phone, zipCode, caseType, description } = formData;

      // Send notification email
      context.log('Preparing notification email');
      const notificationEmail = {
        senderAddress: SENDER_ADDRESS,
        content: {
          subject: `New Contact Form - ${caseType} Case`,
          plainText: `New Contact Form Submission\nCase Type: ${caseType}\nName: ${firstName} ${lastName}\nEmail: ${email}\nPhone: ${phone}\nZip Code: ${zipCode}${description ? `\nDescription: ${description}` : ''}`,
          html: `
            <h2>New Contact Form Submission</h2>
            <p><strong>Case Type:</strong> ${caseType}</p>
            <p><strong>Name:</strong> ${firstName} ${lastName}</p>
            <p><strong>Email:</strong> ${email}</p>
            <p><strong>Phone:</strong> ${phone}</p>
            <p><strong>Zip Code:</strong> ${zipCode}</p>
            ${description ? `<p><strong>Description:</strong><br>${description}</p>` : ''}
          `
        },
        recipients: { 
          to: RECIPIENT_ADDRESSES.map(address => ({ address }))
        }
      };

      context.log('Sending notification email');
      try {
        const notificationPoller = await emailClient.beginSend(notificationEmail);
        context.log('Notification email send initiated, waiting for completion');
        await notificationPoller.pollUntilDone();
        context.log('Notification email sent successfully');
      } catch (emailError: any) {
        context.error('Error sending notification email:', emailError);
        throw new Error(`Failed to send notification email: ${emailError.message}`);
      }

      // Send auto-reply
      context.log('Preparing auto-reply email');
      const autoReplyEmail = {
        senderAddress: SENDER_ADDRESS,
        content: {
          subject: "Thank you for contacting Bull Attorneys",
          plainText: `Thank you for contacting Bull Attorneys\n\nDear ${firstName} ${lastName},\n\nWe have received your inquiry regarding your ${caseType} case. One of our attorneys will review your information and contact you shortly.\n\nIf you need immediate assistance, please call us at (316) 684-4400.\n\nBest regards,\nBull Attorneys Team`,
          html: `
            <h2>Thank you for contacting Bull Attorneys</h2>
            <p>Dear ${firstName} ${lastName},</p>
            <p>We have received your inquiry regarding your ${caseType} case. One of our attorneys will review your information and contact you shortly.</p>
            <p>If you need immediate assistance, please call us at <a href="tel:+13166844400">(316) 684-4400</a>.</p>
            <p>Best regards,<br>Bull Attorneys Team</p>
          `
        },
        recipients: { to: [{ address: email }] }
      };

      context.log('Sending auto-reply email');
      try {
        const autoReplyPoller = await emailClient.beginSend(autoReplyEmail);
        context.log('Auto-reply email send initiated, waiting for completion');
        await autoReplyPoller.pollUntilDone();
        context.log('Auto-reply email sent successfully');
      } catch (emailError: any) {
        context.error('Error sending auto-reply email:', emailError);
        throw new Error(`Failed to send auto-reply email: ${emailError.message}`);
      }

      context.log('Contact form processing completed successfully');
      return {
        status: 200,
        jsonBody: { message: "Form submitted successfully" }
      };
    } catch (emailClientError: any) {
      context.error('Error initializing email client:', emailClientError);
      throw new Error(`Failed to initialize email client: ${emailClientError.message}`);
    }
  } catch (error) {
    context.error('Error processing contact form:', error);
    if (error instanceof Error) {
      context.log(`Error name: ${error.name}`);
      context.log(`Error message: ${error.message}`);
      context.log(`Error stack: ${error.stack}`);
    } else {
      context.log(`Unknown error type: ${typeof error}`);
    }
    
    return {
      status: 500,
      jsonBody: { 
        error: "An error occurred while processing your request",
        message: error instanceof Error ? error.message : String(error)
      }
    };
  }
}

app.http('contactForm', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: contactForm
});
