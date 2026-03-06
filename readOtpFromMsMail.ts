import type { WalnutContext } from './walnut';
import * as https from 'https';

/** @walnut_method
 * name: Read OTP from MS Mail
 * description: Read OTP from MS Mail using Graph API and save to variable
 * actionType: custom_read_otp_ms_mail
 * context: shared
 * needsLocator: false
 * category: Email Automation
 */
export async function readOtpFromMsMail(ctx: WalnutContext) {
  const clientId = '958ef5ed-0ddd-4f7c-b41c-ad1c3572a21e';
  const clientSecret = 'a33fa700-774f-482a-a620-74d2014fd25c';
  const tenantId = 'faa3f9fc-9b37-406b-b37d-58a9f045c17a';
  const emailAddress = 'shreya.g@simplify3x.com';
  const subjectLine = 'Verify your identity in Salesforce';
  
  ctx.log('Getting access token for Microsoft Graph API...');
  
  // Get access token
  const tokenResponse = await makeRequest(
    'POST',
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    `client_id=${clientId}&scope=https://graph.microsoft.com/.default&client_secret=${clientSecret}&grant_type=client_credentials`,
    { 'Content-Type': 'application/x-www-form-urlencoded' }
  );
  
  const accessToken = JSON.parse(tokenResponse).access_token;
  ctx.log('Access token obtained successfully');
  
  // Wait up to 60 seconds for email to arrive
  const maxAttempts = 12; // 12 attempts * 5 seconds = 60 seconds
  let emailFound = false;
  let otpCode = '';
  
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    ctx.log(`Searching for email (attempt ${attempt}/${maxAttempts})...`);
    
    // Search for emails with subject line
    const messagesResponse = await makeRequest(
      'GET',
      `https://graph.microsoft.com/v1.0/users/${emailAddress}/messages?$filter=subject eq '${subjectLine}'&$orderby=receivedDateTime desc&$top=1`,
      null,
      { 'Authorization': `Bearer ${accessToken}` }
    );
    
    const messages = JSON.parse(messagesResponse);
    
    if (messages.value && messages.value.length > 0) {
      const emailBody = messages.value[0].body.content;
      ctx.log('Email found, extracting OTP...');
      
      // Extract OTP using regex pattern "Verification Code: 661161"
      const otpMatch = emailBody.match(/Verification Code:\s*(\d{6})/i);
      
      if (otpMatch && otpMatch[1]) {
        otpCode = otpMatch[1];
        emailFound = true;
        ctx.log(`OTP extracted successfully: ${otpCode}`);
        break;
      } else {
        ctx.warn('Email found but OTP pattern not matched');
      }
    }
    
    if (attempt < maxAttempts) {
      ctx.log('Email not found, waiting 5 seconds before retry...');
      await sleep(5000);
    }
  }
  
  if (!emailFound || !otpCode) {
    throw new Error(`Failed to find email with subject "${subjectLine}" or extract OTP within 60 seconds`);
  }
  
  // Save OTP to variable context
  ctx.setVariable('otp', otpCode);
  ctx.log(`OTP saved to variable 'otp': ${otpCode}`);
}

// Helper function to make HTTPS requests
function makeRequest(method: string, url: string, body: string | null, headers: Record<string, string>): Promise<string> {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: method,
      headers: headers
    };
    
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) {
          resolve(data);
        } else {
          reject(new Error(`HTTP ${res.statusCode}: ${data}`));
        }
      });
    });
    
    req.on('error', reject);
    
    if (body) {
      req.write(body);
    }
    
    req.end();
  });
}

// Helper function to sleep
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}
