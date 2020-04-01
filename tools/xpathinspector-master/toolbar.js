"use strict";

const browser = chrome || browser;
const textarea_expression = document.getElementById("xpath_expression");
const textarea_result = document.getElementById("xpath_result");
const button_run = document.getElementById("run_xpath");
const button_hide = document.getElementById("hide");

function messageCallback(message, sender, sendResponse) 
{
    console.log(`toolbar: messageCallback() message = ${JSON.stringify(message)}`);
   
    if (message.operation === "update")
    {
        textarea_expression.value = message.expression;
        textarea_result.value = message.content;
    }

    sendResponse();
    return true;
}

function init()
{
    console.log("toolbar: enter init...");

    browser.runtime.onMessage.addListener(messageCallback);
    
    button_run.onclick = ()=>{
       if(textarea_expression.value.length > 0)
       {
        browser.runtime.sendMessage({"operation": "forward_request_run_expression", "expression": textarea_expression.value});
       }
    };

    button_hide.onclick = ()=>{
        browser.runtime.sendMessage({"operation": "forward_hide_current_toolbar"});
    }
   
    button_run.innerText = browser.i18n.getMessage("Run");
    button_hide.innerText = browser.i18n.getMessage("CloseToolbar");
    console.log("toolbar: leave init...");
}

init();
// window.addEventListener("load", function(){ init(); });