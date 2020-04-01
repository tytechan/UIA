"use strict";
let toolbar = null;
let inspector = null;
let current_element = null;
const browser = chrome || browser;

class ToolBar
{
	constructor()
	{
		this.frame_ = document.createElement("iframe");
		this.frame_.classList.add("hidden");
		this.frame_.id = "xpath_inspector_toolbar";
		this.frame_.src = browser.runtime.getURL("toolbar.html");
		document.documentElement.appendChild(this.frame_);	
	}

	update(result)
	{
		// 在Chrome，Firefox中，对于Content Script，二者使用runtime.sendMessage时，表现形式并不一样：
		// 在chrome中，所有的扩展页面，都能收到Content Script发送的消息
		// 在Firefox中，只有Background.js能收到消息
		// 所以，取最小合集，使用Background.js接收消息后，再转发给本页面的toolbar.js
		// 备选方案：html5的postMessage也可解决页面通信问题，因为担心对目标页面有意外影响，所以并未采用
		browser.runtime.sendMessage({"operation": "forward_update", "expression": result.expression, "content": result.content});
	}

	// 运行输入的xpath表达式
	runExpression(expression)
	{
		let element_content = "";
		let elements = new Array();
		let result = inspector.evaluate(expression);
		for(let index = 0; result != null && index < result.snapshotLength; ++index)
		{
			elements.push(result.snapshotItem(index));
			element_content += getElementText(result.snapshotItem(index)) + "\n";
		}

		if(elements.length === 0)
		{
			// 未找到任何元素后，在xpath结果区域，给出问题提示
			element_content = browser.i18n.getMessage("NotFindElement");
		}

		this.showHighlightElements(elements);
		this.update({"expression": expression, "content": element_content});
		if(elements.length >= 1)
		{
			// 运行输入的xpath表达式后，找到的元素可能在当前页面中看不到，需要滑动滚动条，自动定位到目标元素
			this.moveToElement(elements[0]);
		}
	}

	switch(enable_inspector)
	{
		if(enable_inspector)
		{
			this.frame_.classList.remove("hidden");
			// document.addEventListener("mouseover", handleInspectElement);
			document.addEventListener("mousedown", handleInspectElement);
		}
		else
		{
			this.frame_.classList.add("hidden");
			// document.removeEventListener("mouseover", handleInspectElement);
			document.removeEventListener("mousedown", handleInspectElement);
		}
	}

	hide()
	{
		if(!this.frame_.classList.contains("hidden"))
		{
			this.frame_.classList.add("hidden");
		}
	}

	showHighlightElement(element)
	{
		this.clearHighlight();
		element.classList.add("xpath_inspector_highlight");
	}

	showHighlightElements(elements)
	{
		// 先清除已显示的高亮元素，再显示当前已选中的元素
		this.clearHighlight();
		for(let index = 0; index < elements.length; ++index)
		{
			elements[index].classList.add("xpath_inspector_highlight");
		}
	}

	clearHighlight()
	{
		let elements = document.querySelectorAll(".xpath_inspector_highlight");
		for(let index = 0; index < elements.length; ++index)
		{
			elements[index].classList.remove("xpath_inspector_highlight");
		}
	}

	moveToElement(element) 
	{
		let rect = element.getBoundingClientRect();
		let top = window.pageYOffset + rect.top;
		let currentTop = 0;
		let requestId;
		
		function step(timestamp) 
		{
		  currentTop += 10;
		  if(currentTop <= top)
		  {
			window.scrollTo(0, currentTop);
			requestId = window.requestAnimationFrame(step);
		  }
		  else
		  {
			window.cancelAnimationFrame(requestId);
		  }
		}
		window.requestAnimationFrame(step);
	  }
}

class Inspector
{
	constructor(){}

	// https://developer.mozilla.org/zh-CN/docs/Web/API/Node
	getXPath(element)
	{
		console.log("getXPath() enter...");
		
		let result = {
			"error": 0, 
			"msg": "ok", 
			"expression": "",
			"url": window.location.href};
		
		if(element === null)
		{
			result["error"] = 1;
			result["msg"] = "element is null";
			return result;
		}

		let content = getElementText(element).trim();
		result["content"] = content;

		let expression = "";
		let tag_name = element.tagName.toLowerCase();
		if(element.id)
		{
			expression = `//${tag_name}[@id="${element.id}"]`;
			if(this.isValidResult(element, this.evaluate(expression)))
			{
				result["expression"]= expression;
				return result;
			}
			else
			{
				console.error(`id: no valid expression, id = ${element.id}`);
			}
		}

		let conditions = new Array();
		if(element.className)
		{
			let valid_classname = false;
			for(let index = 0; index < element.classList.length; ++index)
			{
				if(this.isValidExpression(`//${tag_name}[contains(@class,"${element.classList[index]}")]`, element))
				{
					conditions.push(`contains(@class,"${element.classList[index]}")`);
					valid_classname = true;
				}
			}

			if (!valid_classname) 
			{
				console.error(`className: no valid expression, className = ${element.className}`);
			}
			// if(this.isValidExpression(`//${tag_name}[@class="${element.className}"]`, element))
			// {
			// 	conditions.push(`@class="${element.className}"`);
			// }
			// else
			// {
			// 	console.error(`className: no valid expression, className = ${element.className}`);
			// }
		}

		if(content)
		{
			if(this.isValidExpression(`//${tag_name}[.="${content}"]`, element))
			{
				conditions.push(`.="${content}"`);
			}
			else if(this.isValidExpression(`//${tag_name}[contains(@value,"${content}")]`, element))
			{
				conditions.push(`contains(@value,"${content}")`);
			}
			else if(this.isValidExpression(`//${tag_name}[contains(text(),"${content}")]`, element))
			{
				conditions.push(`contains(text(),"${content}")`);
			}
			else
			{
				console.error(`content: no valid expression, content = ${content}`);
			}
		}

		if(element.name)
		{
			if(this.isValidExpression(`//${tag_name}[@name="${element.name}"]`, element))
			{
				conditions.push(`@name="${element.name}"`);
			}
			else
			{
				console.error(`name: no valid expression, name = ${element.name}`);
			}
		}

		console.log("conditions = " + JSON.stringify(conditions));
		let expressions = this.getCombination(conditions);
		// expressions中，存放的是各个子条件，需要使用and组合起来
		for(let index = 0; index < expressions.length; ++index)
		{
			expression = `//${tag_name}[${expressions[index].join(" and ")}]`;
			if(this.isValidResult(element, this.evaluate(expression)))
			{
				// 只找到一个元素，说明当前表达式满足唯一性要求
				result["expression"]= expression;
				return result;
			}
			else
			{
				// 说明当前表达式不满足要求，继续循环
			}
		}

		// 如果id,class,text,name等组合方式依然无法保证唯一性，那么使用目标元素的其它属性尝试找到元素
		let nodeName = "";
		let attribute_conditions = new Array();
		for(let index = 0; index < element.attributes.length; ++index)
		{
			nodeName = element.attributes[index].nodeName;
			// 记录非id,class,name等属性
			if(nodeName !== "id" && nodeName !== "class" && nodeName !== "name" && nodeName.length > 0 )
			{
				// 只有验证通过的子条件，才能加入数组，为下一步组合条件做准备
				if(this.isValidExpression(`//${tag_name}[@${element.attributes[index].nodeName}="${element.attributes[index].nodeValue}"]`, element))
				{
					attribute_conditions.push(`@${element.attributes[index].nodeName}="${element.attributes[index].nodeValue}"`);
				}
				else
				{
					console.error(`attribute name: no valid expression, attribute name = ${nodeName}`);
				}
			}
		}

		// 使用class，name，text全组合 外加 元素其它属性的方式，查找元素		
		let strict_expression = conditions.join(" and ");	// 使用优先级条件中的全部组合，提高准确度
		let attribute_expressions = this.getCombination(attribute_conditions);
		for(let index = 0; index < attribute_expressions.length; ++index)
		{
			if(strict_expression.length > 0)
			{
				expression = `//${tag_name}[${strict_expression} and ${attribute_expressions[index].join(" and ")}]`;
			}
			else
			{
				expression = `//${tag_name}[${attribute_expressions[index].join(" and ")}]`;
			}
		
			if(this.isValidResult(element, this.evaluate(expression)))
			{
				// 只找到一个元素，说明当前表达式满足唯一性要求
				result["expression"] = expression;
				return result;
			}
			else
			{
				// 说明当前表达式不满足要求，继续循环
			}
		}

		// 最后一步：绝对路径定位 = 标签 + 位置
		expression = "";
		let element_iterator = element;
		while(element_iterator != null && element_iterator.nodeType === Node.ELEMENT_NODE)
		{
			let component = `${element_iterator.tagName.toLowerCase()}`;
			if(element_iterator.id)
			{
				component += `[@id="${element_iterator.id}"]`;
			}
			else
			{
				let element_index = this.getElementIndex(element_iterator);
				if(element_index >= 1)
				{
					component += `[${element_index}]`;
				}
			}

			expression = `/${component}${expression}`
			element_iterator = element_iterator.parentNode || element_iterator.parentElement;
		}
		
		console.log(`absolute path = ${expression}`);
		if(this.isValidResult(element, this.evaluate(expression)))
		{
			// 只找到一个元素，说明当前表达式满足唯一性要求
			result["expression"] = expression;
			return result;
		}
		else
		{
			// 绝对路径也无法找到对应元素，需要人工调整
			console.error("absolute path fail...");
		}

		return result;
	}

	isValidResult(element, xpath_result)
	{
		return xpath_result && xpath_result.snapshotLength === 1 && (xpath_result.snapshotItem(0) === element);
	}

	// 获取当前节点在兄弟节点中的位置
	getElementIndex(element)
	{
		// xpath起始节点的序号为1，不是0，逆序查找兄弟节点
		let index = 1;
		for(let sibling_element = element.previousSibling; sibling_element != null; sibling_element = sibling_element.previousSibling)
		{
			if(sibling_element.nodeType === Node.ELEMENT_NODE && this.isSiblingElement(element, sibling_element))
			{
				++index;
			}
		}

		if(index > 1)
		{
			return index;
		}

		// 判断节点是不是第一个节点，正向查找
		for(let sibling_element = element.nextSibling; sibling_element != null; sibling_element = sibling_element.nextSibling)
		{
			if(sibling_element.nodeType === Node.ELEMENT_NODE && this.isSiblingElement(element, sibling_element))
			{
				return 1;
			}
		}

		return 0;
	}

	isSiblingElement(element, sibling_element)
	{
		return (element.tagName === sibling_element.tagName);
		// return (element.tagName === sibling_element.tagName 
		// 	&& (!element.className || element.className === sibling_element.className) 
		// 	&& (!element.id || element.id === sibling_element.id));
	}

	// https://developer.mozilla.org/zh-CN/docs/Web/API/Document/evaluate
	// https://developer.mozilla.org/en-US/docs/Web/API/XPathResult
	evaluate(expression)
	{
		let result = null;
		try 
		{
			result = document.evaluate(expression, document, null, XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE, null);
		} 
		catch (e) 
		{
			console.log("xpath exception = " + e + " expression = " + expression);
			return null;
		}

		// let log_info = "";
		// let snapshot_element = null;
		// for(let index = 0; index < result.snapshotLength; ++index)
		// {
		// 	snapshot_element = result.snapshotItem(index);
		// 	log_info += `\n### index = ${index} tag name = ${snapshot_element.tagName.toLowerCase()} content = ${snapshot_element.textContent}`;
		// }
		// console.log(`evaluate() expression = ${expression} count = ${result.snapshotLength} ${log_info}`);
		return result;
	}

	isValidExpression(expression, element)
	{
		let result = null;
		try 
		{
			result = document.evaluate(expression, document, null, XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE, null);
		} 
		catch (e) 
		{
			console.log("xpath exception = " + e + " expression = " + expression);
			return false;
		}

		console.log(`###expression = ${expression} count = ${result.snapshotLength}`);
		if(result == null || result.snapshotLength === 0)
		{
			return false;
		}

		for(let index = 0; index < result.snapshotLength; ++index)
		{
			if(result.snapshotItem(index) === element)
			{
				return true;
			}
		}

		return false;
	}

	// 返回数组元素的全部组合
	getCombination(data) 
	{
		let result = [];
		for (let index = 0; index < data.length; ++index) 
		{
			let current_array = [];
			let current = data[index];
			for(let pos = 0; pos < result.length; ++pos)
			{
				current_array.push(result[pos].concat([current]));
			}

			result.push([current]);
			if(current_array.length > 0)
			{
				result = result.concat(current_array);
			}
		}
		
		// 按照数组长度大小排序
		result.sort((first, second)=>{
			return first.length - second.length;
		});

		return result;
	}
}



function handleKeyDown(event)
{
	let ctrlKey = event.ctrlKey || event.metaKey;
	if(!event.shiftKey && !event.altKey && ctrlKey && event.key.toLowerCase() === "enter")
	{
		console.log("键盘输入:Ctrl + Enter");		
		browser.runtime.sendMessage({"operation": "switch_inspector"});
	}
}

function handleInspectElement(event)
{
	if(current_element == event.target)
	{
		return;
	}
	else
	{
		current_element = event.target;
	}

	
	// 用户按下鼠标左键，表示正在探查元素
	if(event.button == 0)
	// if(event.shiftKey)
	{
		let result = inspector.getXPath(current_element);
		toolbar.showHighlightElement(current_element);
		toolbar.update(result);
		postHttp(result);
		console.log(`getXPath() result = ${JSON.stringify(result)}`);
	}
}

function getElementText(element)
{
	if(element == null)
	{
		return "";
	}

	if(element.textContent)
	{
		return element.textContent;
	}

	if(element.innerText)
	{
		return element.innerText;
	}

	if(element.value)
	{
		return element.value;
	}

	if(element.nodeValue)
	{
		return element.nodeValue;
	}

	// if(element.innerHTML)
	// {
	// 	return element.innerHTML;
	// }

	return "";
}

function postHttp(data)
{
	let xhr = new XMLHttpRequest();
	xhr.open("POST", "http://localhost:54321/rpa/xpath", true);
	xhr.responseType = "json";
	xhr.setRequestHeader("Content-Type","application/json");
	xhr.onreadystatechange = function(){
		if(xhr.readyState == 4)
		{
			console.log(`http status = ${xhr.status} response = ${xhr.response}`);
		}
	}
	xhr.send(JSON.stringify(data));
}

function messageCallback(message, sender, sendResponse) 
{
    console.log("inject: messageCallback() message = " + JSON.stringify(message));
    if (message.operation === "update_inspector_status")
    {
		// 收到更新探测器状态通知
		setTimeout(() => { toolbar.switch(message.enable_inspector); }, 1);
	}
	else if(message.operation === "hide_current_toolbar")
	{
		// 隐藏当前页面的探测器工具条
		setTimeout(() => { toolbar.hide() }, 1);
	}
	else if(message.operation === "request_run_expression")
	{
		// 运行人工输入的表达式
		setTimeout(() => {toolbar.runExpression(message.expression);}, 1);
	}

	sendResponse();
    return true;
}

function init() 
{
	console.log("inject: enter init...");

	toolbar = new ToolBar();
	inspector = new Inspector();
	document.addEventListener("keydown", handleKeyDown);
	browser.runtime.onMessage.addListener(messageCallback);
	browser.runtime.sendMessage({"operation": "read_inspector_status"}, (response)=>{
		// 从background页面读取探测器的最新状态
		toolbar.switch(response && response.enable_inspector);
	});
	
    console.log("inject: leave init...");
}

init();