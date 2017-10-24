/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

declare let fabric: any;

(() => {
    "use strict";

    let cellToHighlight: Excel.Range;
    let messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = () => {
        $(document).ready(() => {
            // Initialize the FabricUI notification mechanism and hide it
            let element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            initAzquo();

/*
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
                $("#template-description").text("This sample will display the value of the cells you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");

            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightHighestValue);
            */
        });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, text): void {
        document.getElementById("test").innerHTML = text;
     }

    function initAzquo() {
        Excel.run(function (ctx) {
            var nameditems = ctx.workbook.names;
            nameditems.load('items');
            return ctx.sync().then(function () {
                for (var i = 0; i < nameditems.items.length; i++) {
                    var $rangeName = nameditems.items[i].name;
                    Office.context.document.bindings.addFromNamedItemAsync($rangeName, Office.BindingType.Matrix, { id: $rangeName }, function (result) {
                        if (result.status == Office.AsyncResultStatus.Failed) {
                            showNotification("Error", 'Error trying to bind : ' + $rangeName + ":" + result.error.message);
                        }
                    });
                }
                return ctx.sync()
                    .then(function () {
                        listBindings(ctx);

                    });
            });
        }).catch(function (error) {
            console.log('error: ' + error.text);
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });



        function listBindings(ctx) {
            Office.context.document.bindings.getAllAsync(function (bindings) {
                var bindingString = '';
                for (var i in bindings.value) {
                    Office.select("bindings#" + bindings.value[i].id).addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                    bindingString += bindings.value[i].id + '\n';
                }
                showNotification("Error", 'Existing bindings: ' + bindingString);
                populateBook(ctx,bindings);
            });
            ctx.workbook.worksheets.getActiveWorksheet().getRange("A1").values = [["I as clicked!"]];
            return ctx.sync();
        }


        function onBindingDataChanged(edcea: Excel.BindingDataChangedEventArgs) {
            showNotification("Change", "data changed in " + edcea.binding.id);
            return;
        }


        function populateBook(ctx,bindings: Office.AsyncResult) {
            interface IDictionary {
                [index: string]: string[][];
            }
     

            var bindingString = '';
            for (var i in bindings.value) {
                var rangeName: String = bindings.value[i].id;
                if (rangeName.length > 13 && rangeName.substring(0, 13).toLowerCase() === 'az_dataregion') {
                    var facetArray = {} as IDictionary;

                    //extract all the names that refer to the region, and make text arrays of those names
                    var region: String = rangeName.substring("az_dataregion".length);
                    
                    for (var range in bindings.value) {
                         var binding: Excel.Binding = bindings.value[range];
                         var bind: String = binding.id;
                       if (bind.length > region.length) {
                            var facet: string = bind.substring(0,bind.length - region.length).toLowerCase();
                            if (bind.substring(facet.length) === region) {
                                facetArray[facet] = binding.getRange().text;
   
                            }
                        }
                    }
                    var j = 1;

                    //now get the values
                }
            }
        }
    }


})();
