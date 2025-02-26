(function() {
  // Get cookies as a single string
  const cookies = document.cookie;
  // Split into individual cookie strings
  const cookieArray = cookies.split('; ');
  // Build an object of cookie name-value pairs
  const cookieObject = {};
  cookieArray.forEach(cookie => {
    const [name, ...rest] = cookie.split('=');
    const value = rest.join('=');
    cookieObject[name] = decodeURIComponent(value);
  });
  // Log the JSON string
  console.log(JSON.stringify(cookieObject, null, 2));
})();
