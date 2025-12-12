import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { EmailClient } from "@azure/communication-email";

// Environment variables - set these in Azure Function App settings
const CONNECTION_STRING = process.env.ACS_CONNECTION_STRING as string;
const SENDER_ADDRESS = process.env.SENDER_EMAIL_ADDRESS || "DoNotReply@2dde48cf-f3cb-436b-838a-1c27aa0e1c0c.azurecomm.net";
const RECIPIENT_ADDRESSES = ["brad@bullattorneys.com", "sudheer@bullattorneys.com", "web@bullattorneys.com"];
const RECAPTCHA_SECRET_KEY = process.env.RECAPTCHA_SECRET_KEY as string;

// Rate limiting store (in production, use Redis or a database)
const rateLimitStore = new Map<string, { count: number; resetTime: number }>();

// Blocked emails (all lowercase for case-insensitive matching)
const BLOCKED_EMAILS = [
  'jacobroyvisser55@gmail.com',
  'jacobvisser45@gmail.com',
  // Add more blocked emails here
];

interface ContactFormData {
  firstName: string;
  lastName: string;
  email: string;
  phone: string;
  zipCode: string;
  caseType: string;
  description?: string;
  recaptchaToken?: string;
  website?: string; // Honeypot field
}

// Rate limiting function
function checkRateLimit(identifier: string, maxRequests = 3, windowMs = 60000): boolean {
  const now = Date.now();
  const record = rateLimitStore.get(identifier);

  if (!record || now > record.resetTime) {
    rateLimitStore.set(identifier, { count: 1, resetTime: now + windowMs });
    return true;
  }

  if (record.count >= maxRequests) {
    return false;
  }

  record.count++;
  return true;
}

// reCAPTCHA validation function
async function validateRecaptcha(token: string): Promise<{ success: boolean; score: number }> {
  try {
    const response = await fetch('https://www.google.com/recaptcha/api/siteverify', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: `secret=${RECAPTCHA_SECRET_KEY}&response=${token}`,
    });

    const data = await response.json();
    return {
      success: data.success || false,
      score: data.score || 0
    };
  } catch (error) {
    console.error('reCAPTCHA validation error:', error);
    return { success: false, score: 0 };
  }
}

// Gibberish detection function
function isGibberish(text: string): boolean {
  if (!text || text.length < 2) return false;
  
  // Check for vowels - real names usually have vowels
  const vowelCount = (text.match(/[aeiouAEIOU]/g) || []).length;
  const consonantCount = text.length - vowelCount;
  
  // If no vowels or too many consonants in a row, likely gibberish
  if (vowelCount === 0 || consonantCount / text.length > 0.8) {
    return true;
  }
  
  // Check for excessive repeated characters
  const repeatedChars = text.match(/(.)\1{2,}/g);
  if (repeatedChars && repeatedChars.length > 0) {
    return true;
  }
  
  return false;
}

// URL spam detection
function containsUrls(text: string): boolean {
  const urlRegex = /(https?:\/\/[^\s]+|www\.[^\s]+|[^\s]+\.(com|net|org|edu|gov|mil|int|co|io|ly|me|tv|info|biz|name|mobi|tel|travel)[^\s]*)/gi;
  return urlRegex.test(text);
}

// Email validation
function isValidEmail(email: string): boolean {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// Phone validation
function isValidPhone(phone: string): boolean {
  const phoneDigits = phone.replace(/\D/g, '');
  return phoneDigits.length >= 10;
}

// Zip code validation
function isValidZipCode(zipCode: string): boolean {
  const zipRegex = /^\d{5}(-\d{4})?$/;
  return zipRegex.test(zipCode);
}

export async function contactForm(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  context.log('Contact form function triggered');

  // Handle CORS preflight
  if (request.method === 'OPTIONS') {
    return {
      status: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Accept',
        'Access-Control-Max-Age': '86400'
      }
    };
  }

  // Only allow POST requests
  if (request.method !== 'POST') {
    return {
      status: 405,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    // Get client IP for rate limiting
    const clientIp = request.headers.get('x-forwarded-for') || 
                     request.headers.get('x-real-ip') || 
                     'unknown';

    // Parse request body
    let formData: ContactFormData;
    try {
      const body = await request.text();
      formData = JSON.parse(body);
    } catch (parseError) {
      context.log('Invalid JSON in request body:', parseError);
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Invalid request format' })
      };
    }

    context.log('Received form data:', { 
      firstName: formData.firstName, 
      lastName: formData.lastName, 
      email: formData.email,
      caseType: formData.caseType,
      hasRecaptcha: !!formData.recaptchaToken 
    });

    // Honeypot check
    if (formData.website && formData.website.trim() !== '') {
      context.log('Honeypot triggered - bot detected');
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Spam detected' })
      };
    }

    // Check blocked emails
    if (formData.email && BLOCKED_EMAILS.includes(formData.email.toLowerCase())) {
      context.log('Blocked email attempted:', formData.email);
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Email not allowed' })
      };
    }

    // Rate limiting
    if (!checkRateLimit(clientIp, 3, 60000)) {
      context.log('Rate limit exceeded for IP:', clientIp);
      return {
        status: 429,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Too many requests. Please try again later.' })
      };
    }

    // Validate required fields
    const requiredFields = ['firstName', 'lastName', 'email', 'phone', 'zipCode', 'caseType'];
    for (const field of requiredFields) {
      if (!formData[field as keyof ContactFormData] || 
          String(formData[field as keyof ContactFormData]).trim() === '') {
        return {
          status: 400,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ error: `${field} is required` })
        };
      }
    }

    // Validate field lengths
    const maxLengths = {
      firstName: 50,
      lastName: 50,
      email: 100,
      phone: 20,
      zipCode: 10,
      caseType: 50,
      description: 1200
    };

    for (const [field, maxLength] of Object.entries(maxLengths)) {
      const value = formData[field as keyof ContactFormData];
      if (value && String(value).length > maxLength) {
        return {
          status: 400,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ error: `${field} is too long (max ${maxLength} characters)` })
        };
      }
    }

    // Validate email format
    if (!isValidEmail(formData.email)) {
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Invalid email format' })
      };
    }

    // Validate phone number
    if (!isValidPhone(formData.phone)) {
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Invalid phone number' })
      };
    }

    // Validate zip code
    if (!isValidZipCode(formData.zipCode)) {
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Invalid zip code format' })
      };
    }

    // Check for gibberish in names
    if (isGibberish(formData.firstName) || isGibberish(formData.lastName)) {
      context.log('Gibberish detected in names');
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Please enter valid names' })
      };
    }

    // Check for URLs in description (spam protection)
    if (formData.description && containsUrls(formData.description)) {
      context.log('URLs detected in description - potential spam');
      return {
        status: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'URLs are not allowed in the description' })
      };
    }

    // Validate reCAPTCHA if token is provided
    if (formData.recaptchaToken) {
      if (!RECAPTCHA_SECRET_KEY) {
        context.log('reCAPTCHA secret key not configured');
        return {
          status: 500,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ error: 'reCAPTCHA not configured' })
        };
      }

      const recaptchaResult = await validateRecaptcha(formData.recaptchaToken);
      if (!recaptchaResult.success || recaptchaResult.score < 0.5) {
        context.log('reCAPTCHA validation failed:', recaptchaResult);
        return {
          status: 400,
          headers: {
            'Access-Control-Allow-Origin': '*',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ error: 'reCAPTCHA validation failed' })
        };
      }
      context.log('reCAPTCHA validation successful, score:', recaptchaResult.score);
    }

    // Send email using Azure Communication Services
    context.log('Checking CONNECTION_STRING:', CONNECTION_STRING ? 'Present' : 'Missing');
    context.log('SENDER_ADDRESS:', SENDER_ADDRESS);
    context.log('RECIPIENT_ADDRESSES:', RECIPIENT_ADDRESSES);
    
    if (!CONNECTION_STRING) {
      context.log('Azure Communication Services connection string not configured');
      return {
        status: 500,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ error: 'Email service not configured' })
      };
    }

    context.log('Creating EmailClient...');
    const emailClient = new EmailClient(CONNECTION_STRING);
    context.log('EmailClient created successfully');

    // Create email content
    const emailSubject = `New Contact Form Submission - ${formData.caseType}`;
    const emailBody = `
      <h2>New Contact Form Submission</h2>
      <p><strong>Name:</strong> ${formData.firstName} ${formData.lastName}</p>
      <p><strong>Email:</strong> ${formData.email}</p>
      <p><strong>Phone:</strong> ${formData.phone}</p>
      <p><strong>Zip Code:</strong> ${formData.zipCode}</p>
      <p><strong>Case Type:</strong> ${formData.caseType}</p>
      ${formData.description ? `<p><strong>Description:</strong><br>${formData.description.replace(/\n/g, '<br>')}</p>` : ''}
      <p><strong>Submitted:</strong> ${new Date().toLocaleString('en-US', { timeZone: 'America/Chicago', dateStyle: 'medium', timeStyle: 'medium' })}</p>
      <p><strong>IP Address:</strong> ${clientIp}</p>
    `;

    const emailMessage = {
      senderAddress: SENDER_ADDRESS,
      content: {
        subject: emailSubject,
        html: emailBody,
      },
      recipients: {
        to: RECIPIENT_ADDRESSES.map(email => ({ address: email })),
      },
    };

    try {
      const poller = await emailClient.beginSend(emailMessage);
      const result = await poller.pollUntilDone();
      context.log('Email sent successfully:', result.id);
    } catch (emailError: any) {
      const errorMessage = emailError?.message || String(emailError);
      context.log('Failed to send email:', errorMessage);
      context.error('Email error details:', emailError);
      return {
        status: 500,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ 
          error: 'Failed to send email',
          details: errorMessage 
        })
      };
    }

    // Return success response
    return {
      status: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ 
        success: true, 
        message: 'Form submitted successfully' 
      })
    };

  } catch (error) {
    context.log('Unexpected error:', error);
    return {
      status: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ error: 'Internal server error' })
    };
  }
}

app.http('contactForm', {
    methods: ['POST', 'OPTIONS'],
    authLevel: 'anonymous',
    handler: contactForm
});