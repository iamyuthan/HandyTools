function convertToExtensionJSON(rawCookies) {
  const lines = rawCookies.trim().split('\n');
  const cookies = [];
  
  lines.forEach(line => {
    if (!line.trim()) return;
    
    const cols = line.split('\t');
    
    const name = cols[0]?.trim();
    const value = cols[1]?.trim();
    const domain = cols[2]?.trim();
    const path = cols[3]?.trim() || '/';
    const expires = cols[4]?.trim();
    const httpOnly = cols[6]?.trim() === '✓';
    const secure = cols[7]?.trim() === '✓';
    const sameSiteRaw = cols[8]?.trim() || '';
    
    if (!name) return;
    
    // Convert expiry to Unix timestamp
    let expirationDate = null;
    let session = true;
    
    if (expires && expires !== 'Session') {
      const expiryDate = new Date(expires);
      if (!isNaN(expiryDate.getTime())) {
        expirationDate = Math.floor(expiryDate.getTime() / 1000);
        session = false;
      }
    }
    
    // Convert SameSite to extension format
    let sameSite = null;
    if (sameSiteRaw.toLowerCase() === 'none') {
      sameSite = 'no_restriction';
    } else if (sameSiteRaw.toLowerCase() === 'lax') {
      sameSite = 'lax';
    } else if (sameSiteRaw.toLowerCase() === 'strict') {
      sameSite = 'strict';
    }
    
    // Determine hostOnly (true if domain doesn't start with .)
    const hostOnly = !domain.startsWith('.');
    
    cookies.push({
      domain: domain,
      expirationDate: expirationDate,
      hostOnly: hostOnly,
      httpOnly: httpOnly,
      name: name,
      path: path,
      sameSite: sameSite,
      secure: secure,
      session: session,
      storeId: null,
      value: value
    });
  });
  
  const jsonOutput = JSON.stringify(cookies, null, 2);
  
  // Copy to clipboard
  copy(jsonOutput);
  console.log('📋 Copied to clipboard! Paste into Cookie-Editor Import.');
  console.log(`\n🍪 Total cookies: ${cookies.length}\n`);
  console.log(jsonOutput);
  
  return cookies;
}
