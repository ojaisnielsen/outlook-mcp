#!/usr/bin/env node

console.error('==== DEBUG INFORMATION ====');
console.error('ARGUMENTS:', process.argv);
console.error('ENVIRONMENT VARIABLES:');
// Mask anything that looks like a credential so it never lands in log files
const SENSITIVE = /SECRET|TOKEN|PASSWORD|PASSWD|KEY|CREDENTIAL|AUTH|COOKIE|SESSION/i;
Object.keys(process.env).sort().forEach(key => {
  const value = process.env[key];
  const shown = SENSITIVE.test(key) && value ? `<redacted, ${value.length} chars>` : value;
  console.error(`  ${key}: ${shown}`);
});
console.error('==== END DEBUG INFO ====');

// Load the real application
try {
  require('./index');
} catch (error) {
  console.error('ERROR LOADING INDEX.JS:', error);
} 