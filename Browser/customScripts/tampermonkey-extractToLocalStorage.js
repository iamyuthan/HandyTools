// ==UserScript==
// @name         Extract ID and Store as Token
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Extract 'id' from JSON response and store it as 'token' in local storage
// @author       You
// @match        *://abc.com/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Intercept all fetch requests
    const originalFetch = window.fetch;
    window.fetch = function(...args) {
        return originalFetch(...args).then(response => {
            const clonedResponse = response.clone();

            clonedResponse.json().then(data => {
                if (data.id) {
                    // Store the id in local storage as 'token'
                    localStorage.setItem('token', data.id);
                    console.log(`Token stored: ${data.id}`);
                }
            }).catch(err => {
                console.error('Error parsing JSON:', err);
            });

            return response;
        });
    };
})();
