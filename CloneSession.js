
// Function to export cookies, localStorage, and sessionStorage to Base64 encoded JSON
function exportDataToBase64() {
  // Function to get cookies as an object
  function getCookies() {
    const cookies = document.cookie;
    const cookiesArray = cookies.split('; ');
    const cookiesObj = {};
    cookiesArray.forEach(cookie => {
      const [key, value] = cookie.split('=');
      cookiesObj[key] = value;
    });
    return cookiesObj;
  }

  // Get cookies, localStorage, and sessionStorage
  const cookies = getCookies();
  const localStorageData = { ...localStorage };
  const sessionStorageData = { ...sessionStorage };

  // Create a combined object
  const combinedData = {
    cookies,
    localStorage: localStorageData,
    sessionStorage: sessionStorageData
  };

  // Convert the combined object to a JSON string
  const jsonString = JSON.stringify(combinedData);

  // Encode the JSON string to Base64
  const base64String = btoa(jsonString);

  return base64String;
}

// Call the function and log the result
console.log(exportDataToBase64());


-------------------------------------------------------


// Function to decode Base64 and restore cookies, localStorage, and sessionStorage
function restoreDataFromBase64(base64String) {
  // Decode the Base64 string to JSON string
  const jsonString = atob(base64String);

  // Parse the JSON string to an object
  const dataObj = JSON.parse(jsonString);

  // Restore cookies
  const cookies = dataObj.cookies;
  for (const key in cookies) {
    if (cookies.hasOwnProperty(key)) {
      document.cookie = `${key}=${cookies[key]}; path=/;`;
    }
  }

  // Restore localStorage
  const localStorageData = dataObj.localStorage;
  for (const key in localStorageData) {
    if (localStorageData.hasOwnProperty(key)) {
      localStorage.setItem(key, localStorageData[key]);
    }
  }

  // Restore sessionStorage
  const sessionStorageData = dataObj.sessionStorage;
  for (const key in sessionStorageData) {
    if (sessionStorageData.hasOwnProperty(key)) {
      sessionStorage.setItem(key, sessionStorageData[key]);
    }
  }
}
