// Replace this multiline string with the cookie lines you copied.
// Make sure each cookie is on its own line, and columns are tab-separated.
const rawCookies = `

`.trim();

const lines = rawCookies.split('\n');

// Map each line to a cookie object
const cookies = lines.map(line => {
  // Split on one or more tabs
  const cols = line.split(/\t+/);

  // columns:
  // 0 => Name
  // 1 => Value
  // 2 => Domain
  // 3 => Path
  // 4 => Expires (Session or ISO date)
  // 5 => Size (not needed)
  // 6 => HttpOnly (✓ or blank)
  // 7 => Secure (✓ or blank)
  // 8 => SameSite (None, Lax, Strict, or blank)
  
  return {
    name: cols[0],
    value: cols[1],
    domain: cols[2],
    path: cols[3],
    expires: cols[4],
    httpOnly: cols[6] === '✓',
    secure: cols[7] === '✓',
    sameSite: cols[8] || ''
  };
});

// Print the result as JSON
console.log(JSON.stringify(cookies, null, 2));
