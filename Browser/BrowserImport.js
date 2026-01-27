function importCookiesFromTSV(rawCookies) {
  const lines = rawCookies.trim().split('\n');
  let imported = 0, skipped = 0;
  
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
    const sameSite = cols[8]?.trim() || '';
    
    if (!name || !value) return;
    
    // Build cookie string
    let cookieStr = `${name}=${value}`;
    cookieStr += `; path=${path}`;
    
    // Domain - only set if it matches current domain
    // (Browser will reject cross-domain cookies)
    
    // Expiry
    if (expires && expires !== 'Session') {
      const expiryDate = new Date(expires);
      if (!isNaN(expiryDate.getTime())) {
        cookieStr += `; expires=${expiryDate.toUTCString()}`;
      }
    } else {
      cookieStr += `; max-age=31536000`; // 1 year default
    }
    
    // Secure flag
    if (secure) {
      cookieStr += '; secure';
    }
    
    // SameSite
    if (sameSite) {
      cookieStr += `; samesite=${sameSite}`;
    }
    
    // Set the cookie
    document.cookie = cookieStr;
    
    if (httpOnly) {
      console.log(`⚠️ SET (needs manual HttpOnly): ${name}`);
    } else {
      console.log(`✅ SET: ${name}`);
    }
    imported++;
  });
  
  console.log(`\n📊 Done! Imported: ${imported} cookies`);
  console.log(`💡 Tip: Manually enable HttpOnly for marked cookies in DevTools`);
}
