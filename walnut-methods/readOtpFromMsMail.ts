import type { WalnutContext } from './walnut';
import * as https from 'https';

/** @walnut_method
 * name: Read OTP from MS Mail
 * description: Read OTP from email with subject ${subjectLine} for ${emailAddress} using clientId ${clientId} clientSecret ${clientSecret} tenantId ${tenantId} and store in $[otp]
 * actionType: custom_read_otp_ms_mail
 * context: shared
 * needsLocator: false
 * category: Email Automation
 */
export async function readOtpFromMsMail(ctx: WalnutContext) {
  // Get parameters from step arguments (from ${} and $[] placeholders in description)
  const subjectLine = ctx.args[0];              // from ${subjectLine}
  const emailAddress = ctx.args[1];             // from ${emailAddress}
  const clientId = ctx.args[2];                 // from ${clientId}
  const clientSecret = ctx.args[3];             // from ${clientSecret}
  const tenantId = ctx.args[4];                 // from ${tenantId}
  const outputVar = ctx.args[5];                // from $[otp] — runtime variable name
  
  if (!clientId || !clientSecret || !tenantId) {
    throw new Error('Missing required MS Graph API credentials: clientId, clientSecret, and tenantId');
  }
  
  if (!subjectLine || !emailAddress) {
    throw new Error('Missing required parameters: subjectLine and emailAddress');
  }
  
  ctx.log(`Reading OTP from email: ${emailAddress}, subject: "${subjectLine}"`);
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
    
    // Fetch recent messages without complex filter - get top 10 recent messages
    const messagesResponse = await makeRequest(
      'GET',
      `https://graph.microsoft.com/v1.0/users/${emailAddress}/messages?$top=10&$orderby=receivedDateTime desc`,
      null,
      { 'Authorization': `Bearer ${accessToken}` }
    );
    
    const messages = JSON.parse(messagesResponse);
    
    if (messages.value && messages.value.length > 0) {
      // Search through recent messages for matching subject (partial match)
      for (const message of messages.value) {
        if (message.subject && message.subject.includes(subjectLine)) {
          const emailBody = message.body.content;
          ctx.log(`Email found with subject: "${message.subject}", extracting OTP...`);
          
          // Extract OTP using regex patterns (try multiple common formats)
          const otpPatterns = [
            /Verification Code:\s*(\d{6})/i,
            /OTP:\s*(\d{6})/i,
            /code:\s*(\d{6})/i,
            /\b(\d{6})\b/
          ];
          
          for (const pattern of otpPatterns) {
            const otpMatch = emailBody.match(pattern);
            if (otpMatch && otpMatch[1]) {
              otpCode = otpMatch[1];
              emailFound = true;
              ctx.log(`OTP extracted successfully: ${otpCode}`);
              break;
            }
          }
          
          if (emailFound) {
            break;
          } else {
            ctx.warn('Email found but OTP pattern not matched');
          }
        }
      }
      
      if (emailFound) {
        break;
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
  
  // Save OTP to variable context using the variable name from $[otp] placeholder
  ctx.setVariable(outputVar, otpCode);
  ctx.log(`OTP saved to variable '${outputVar}': ${otpCode}`);
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
