import type { WalnutContext } from './walnut';
import * as https from 'https';

/** @walnut_method
 * name: Read OTP from MS Mail
 * description: Read OTP from email with subject ${subjectLine} for ${emailAddress} using credentials ${clientId} ${clientSecret} ${tenantId} and store in $[otp]
 * actionType: custom_read_otp_ms_mail
 * context: shared
 * needsLocator: false
 * category: Email Automation
 */
export async function readOtpFromMsMail(ctx: WalnutContext) {
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

  // Get all parameters from step arguments (from ${} placeholders in description)
  const subjectLine = ctx.args[0];    // from ${subjectLine}
  const emailAddress = ctx.args[1];   // from ${emailAddress}
  const clientId = ctx.args[2];       // from ${clientId}
  const clientSecret = ctx.args[3];   // from ${clientSecret}
  const tenantId = ctx.args[4];       // from ${tenantId}
  const outputVar = ctx.args[5];      // from $[otp] — runtime variable name
  
  if (!clientId || !clientSecret || !tenantId) {
    throw new Error('Missing required MS Graph API credentials: clientId, clientSecret, and tenantId');
  }
  
  if (!subjectLine || !emailAddress) {
    throw new Error('Missing required parameters: subjectLine and emailAddress');
  }
  
  ctx.log(`Reading OTP from email: ${emailAddress}, subject: "${subjectLine}"`);
  
  // Wait 5 seconds at the beginning to allow email to arrive
  ctx.log('Waiting 5 seconds for email to arrive...');
  await sleep(10000);
  
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
  let messageId = ''; // Store message ID for deletion
  
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    ctx.log(`Searching for email (attempt ${attempt}/${maxAttempts})...`);
    
    // Fetch recent messages - get top 10 recent messages
    const messagesResponse = await makeRequest(
      'GET',
      `https://graph.microsoft.com/v1.0/users/${emailAddress}/messages?$top=10&$orderby=receivedDateTime desc`,
      null,
      { 'Authorization': `Bearer ${accessToken}` }
    );
    
    const messages = JSON.parse(messagesResponse);
    
    if (messages.value && messages.value.length > 0) {
      ctx.log(`Found ${messages.value.length} recent emails, checking each one...`);
      
      // Search through recent messages for matching subject (partial match)
      for (const message of messages.value) {
        const receivedTime = new Date(message.receivedDateTime);
        
        ctx.log(`Checking email: "${message.subject}" received at ${receivedTime.toISOString()}`);
        
        if (message.subject && message.subject.includes(subjectLine)) {
          ctx.log(`  ✓ Subject matches! Looking for OTP...`);
          const emailBody = message.body.content;
          
          // Strip HTML tags and get plain text for better matching
          const plainText = emailBody.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
          ctx.log(`  Email body preview (first 300 chars): ${plainText.substring(0, 300)}`);
          
          // Extract OTP using regex patterns (try multiple common formats)
          const otpPatterns = [
            /(?:verification\s*code|otp|code)[\s:]*(\d{4,8})/i,  // "Verification Code: 123456"
            /\b(\d{6})\b/,                                        // Standalone 6-digit number
            /\b(\d{4})\b/,                                        // Standalone 4-digit number
            /(?:pin|token)[\s:]*(\d{4,8})/i,                     // PIN or token patterns
            /is[\s:]*(\d{4,8})/i                                  // "Your code is: 123456"
          ];
          
          for (let i = 0; i < otpPatterns.length; i++) {
            const pattern = otpPatterns[i];
            const otpMatch = plainText.match(pattern);
            if (otpMatch && otpMatch[1]) {
              otpCode = otpMatch[1].trim();
              emailFound = true;
              messageId = message.id; // Save message ID for deletion
              ctx.log(`  ✓ OTP FOUND: ${otpCode} (using pattern #${i + 1}), Email ID: ${messageId}`);
              break;
            }
          }
          
          if (emailFound) {
            break;
          } else {
            ctx.warn('  ❌ Email found but no OTP pattern matched');
            ctx.warn(`  Full email body: ${plainText.substring(0, 1000)}`);
          }
        } else {
          ctx.log(`  ❌ Subject does not match (looking for: "${subjectLine}")`);
        }
      }
      
      if (emailFound) {
        break;
      }
    }
    
    if (attempt < maxAttempts) {
      ctx.log('Email not found yet, waiting 5 seconds before retry...');
      await sleep(5000);
    }
  }
  
  ctx.log(`DEBUG: emailFound=${emailFound}, otpCode=${otpCode}, messageId=${messageId}`);
  
  if (!emailFound || !otpCode) {
    throw new Error(`Failed to find email with subject "${subjectLine}" or extract OTP within 60 seconds`);
  }
  
  ctx.log(`DEBUG: About to call ctx.setVariable with outputVar="${outputVar}", otpCode="${otpCode}"`);
  
  // Save OTP to variable context using the variable name from $[otp] placeholder
  ctx.setVariable(outputVar, otpCode);
  ctx.log(`✓ OTP saved to variable '${outputVar}': ${otpCode}`);
  
  // Delete the email after successfully reading the OTP
  if (messageId) {
    ctx.log(`Deleting email with ID: ${messageId}...`);
    try {
      await makeRequest(
        'DELETE',
        `https://graph.microsoft.com/v1.0/users/${emailAddress}/messages/${messageId}`,
        null,
        { 'Authorization': `Bearer ${accessToken}` }
      );
      ctx.log('✓ Email deleted successfully');
    } catch (deleteError) {
      ctx.warn(`Failed to delete email: ${deleteError}`);
      // Don't throw error - OTP was already saved successfully
    }
  }
}
