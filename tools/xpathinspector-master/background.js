"use strict";
// Chrome/Firefox扩展API兼容性列表
// https://developer.mozilla.org/en-US/docs/Mozilla/Add-ons/WebExtensions/Browser_support_for_JavaScript_APIs
const browser = chrome || browser;
// 默认状态下，探测器没有启用
let enable_inspector = false;

function sendTabMessage(tabId, message)
{
    try
    {
        browser.tabs.sendMessage(tabId, message, ()=>{
            if(chrome.runtime.lastError)
            {
                console.log(`sendMessage() error = ${chrome.runtime.lastError}, message = ${JSON.stringify(message)}`);
            }
        });
    }
    catch(e)
    {
        console.log(`sendMessage() exception = ${e}, message = ${JSON.stringify(message)}`);
    }
}

function switchInspector()
{
    enable_inspector = !enable_inspector;

    if(enable_inspector)
    {
        browser.browserAction.setIcon({path: browser.runtime.getURL("images/enable_icon48.png")});
        browser.browserAction.setTitle({"title": browser.i18n.getMessage("EnableInspector")});
        browser.notifications.create("", {
            "type": "basic", 
            "iconUrl": browser.runtime.getURL("images/enable_icon128.png"), 
            "title": browser.i18n.getMessage("XPathInspector"), 
            "message": browser.i18n.getMessage("EnableInspector"),
            "isClickable": true,
            "priority": 0,
        });
    }
    else
    {
        // 默认图标为红色，就是不可用
        browser.browserAction.setIcon({path: browser.runtime.getURL("images/icon32.png")});
        browser.browserAction.setTitle({"title": browser.i18n.getMessage("DisableInspector")});
        browser.notifications.create("", {
            "type": "basic", 
            "iconUrl": browser.runtime.getURL("images/icon128.png"), 
            "title": browser.i18n.getMessage("XPathInspector"), 
            "message": browser.i18n.getMessage("DisableInspector"),
            "isClickable": true,
            "priority": 0,
        });
    }

    // 广播所有注入的页面，更新自身的状态，向各个注入脚本发送消息
    // 筛选页面时，需要排除浏览器自身设置等页面，否则，在firefox下，发送消息失败时，firefox会报错
    browser.tabs.query({"url": ["http://*/*", "https://*/*"]}, (tabs)=>{
        for(let index = 0; index < tabs.length; ++index)
        {
            // 排除chrome，firefox的扩展商店
            if(tabs[index].url.indexOf("https://chrome.google.com/webstore") !== 0 && tabs[index].url.indexOf("https://addons.mozilla.org") !== 0)
            {
                sendTabMessage(tabs[index].id, {"operation": "update_inspector_status", "enable_inspector": enable_inspector});
            }
        }
    });
}

function messageCallback(message, sender, sendResponse) 
{
    console.log("messageCallback() message = " + JSON.stringify(message));
    if (message.operation === "switch_inspector")
    {
        switchInspector();
    }
    else if(message.operation === "read_inspector_status")
    {
        sendResponse({"enable_inspector": enable_inspector});
        return true;
    }
    else if(message.operation === "forward_update")
    {
        // 转发消息
        message.operation = "update";
        sendTabMessage(sender.tab.id, message);
    }
    else if(message.operation === "forward_hide_current_toolbar")
    {
        // 转发消息
        message.operation = "hide_current_toolbar";
        sendTabMessage(sender.tab.id, message);
    }
    else if(message.operation === "forward_request_run_expression")
    {
        // 转发消息
        message.operation = "request_run_expression";
        sendTabMessage(sender.tab.id, message);
    }

    sendResponse();
    return true;
}


function init() 
{
    console.log("enter init...");

    browser.runtime.onMessage.addListener(messageCallback);
    browser.browserAction.onClicked.addListener((tab)=>{ switchInspector(); });

    console.log("leave init...");
}

window.addEventListener("load", function(){ init(); });