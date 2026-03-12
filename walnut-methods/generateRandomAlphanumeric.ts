import type { WalnutContext } from './walnut';
import * as crypto from 'crypto';

/** @walnut_method
 * name: Generate Random Alphanumeric
 * description: Generate random alphanumeric string of length ${length} and store in $[randomValue]
 * actionType: custom_generate_random_alphanumeric
 * context: shared
 * needsLocator: false
 * category: Data Processing
 */
export async function generateRandomAlphanumeric(ctx: WalnutContext) {
  // ctx.args[0] contains the length parameter from ${length} placeholder
  // ctx.args[1] contains the variable name from $[randomValue] placeholder
  const length = parseInt(ctx.args[0]) || 10; // Default to 10 if not provided or invalid
  const outputVar = ctx.args[1];
  
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  
  // Generate random alphanumeric string
  const randomBytes = crypto.randomBytes(length);
  for (let i = 0; i < length; i++) {
    result += characters.charAt(randomBytes[i] % characters.length);
  }
  
  ctx.log(`Generated random alphanumeric string: ${result}`);
  
  // Store the generated string in variable context for use in subsequent steps
  ctx.setVariable(outputVar, result);
  ctx.log(`✓ Random value saved to variable '${outputVar}': ${result}`);
  
  return result;
}
