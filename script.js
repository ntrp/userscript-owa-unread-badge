// ==UserScript==
// @name          Display Unread Badge for OWA
// @namespace     ntrp
// @version       1.2
// @description   This will check for unread messages and put the unread count in the Outlook Web App.
// @match         https://outlook.office.com/*
// @require       https://cdnjs.cloudflare.com/ajax/libs/favico.js/0.3.10/favico.min.js
// ==/UserScript==

var favicon = new Favico({
    animation: 'popFade'
});

var unread = 0;
var interval;

function update() {
	try {
        // query spans containing read count for all favorite folders
	    var currentUnread = [...document.querySelectorAll('span[autoId="_n_T"]')]
            // get value
            .map(node => node.innerHTML)
            // filter empty ones
            .filter(val => val.length > 0)
            // transform to int
            .map(val => parseInt(val))
            // sum
            .reduce((a, b) => a + b, 0);
        // if not changed do nothing
        if (currentUnread != unread) {
            unread = currentUnread;
            if (unread <= 9) {
                favicon.badge(unread);
            } else {
                // show plus for more than 9 unreads
                favicon.badge("+");
            }
        }
	} catch(err) {
        // in error log error and cancel the function interval
        console.log(err);
        if (interval) {
            window.clearInterval(interval);
        }
    }
}

interval = window.setInterval(update, 1000);
