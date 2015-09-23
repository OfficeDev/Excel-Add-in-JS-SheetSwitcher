/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */

/// <reference path="../App.js" />

(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			// Add a click event handler for the Add Sheets button
			$("#add-sheets").click(addNewSheets);

			// Add a click event handler for the Refresh button
			$("#refresh-button").click(createSheetButtons);

			// Create buttons in the task pane for each sheet in the workbook
			createSheetButtons();
		});
	}


	function addNewSheets() {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {
			// Create a proxy object for the worksheets collection
			var worksheets = ctx.workbook.worksheets;

			// Add 14 sheets to the workbook
			for (var i = 2; i <= 15; i++) {
				// Queue commands to add new sheets to the workbook
				worksheets.add("Sheet" + i);
			}

			//Disable the button
			$('#add-sheets').prop('disabled', true);


			//Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();

		})
		.then(function () {
			// Now that we have sheets, create buttons for each sheet
			// in the taskpane to enable switching
			createSheetButtons();
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}


	function createSheetButtons() {

		// Remove the existing buttons
		$("#buttons-div").empty();

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {
			// Create a proxy object for the worksheets collection 
			var worksheets = ctx.workbook.worksheets;

			// Queue a command to load the name property of each worksheet
			worksheets.load("name");

			//Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync()
				.then(function () {
					//create a button for each sheet in the task pane
					for (var i = 0; i < worksheets.items.length; i++) {
						var buttonName = worksheets.items[i].name;
						var $input = $('<input type="button" class="ms-Button" value=' + buttonName + '>');
						$input.appendTo($("#buttons-div"));
						// Add a click event handler for the button
						(function (buttonName) {
							$input.click(function (e) {
								makeActiveSheet(buttonName);
							});
						})(buttonName);
					}
				});
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}


	function makeActiveSheet(buttonName) {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {
			// Create a proxy object for the worksheets collection 
			var worksheets = ctx.workbook.worksheets;

			// Queue a command to get the sheet with the name of the clicked button
			var clickedSheet = worksheets.getItem(buttonName);

			// Queue a command to insert the sheet name into a cell for easy viewing
			clickedSheet.getCell(0, 0).values = buttonName;

			// Queue a command to activate the clicked sheet
			clickedSheet.activate();

			//Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}
})();

/*
Excel-Add-in-JS-SheetSwitcher, https://github.com/OfficeDev/Excel-Add-in-JS-SheetSwitcher

Copyright (c) Microsoft Corporation

All rights reserved. 

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
associated documentation files (the "Software"), to deal in the Software without restriction, including 
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial 
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT 
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT 
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN 
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE 
USE OR OTHER DEALINGS IN THE SOFTWARE. 
*/