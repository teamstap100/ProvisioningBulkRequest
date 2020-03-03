'use strict';

(function () {
    var itAdminsApiUrl = "https://tap-validation-tab-admin-2.azurewebsites.net/api/tenants/";

    var spinner = '<i class="fa fa-spinner fa-spin"></i>  ';

    microsoftTeams.initialize();

    jQuery.extend(jQuery.expr[':'], {
        invalid: function (elem, index, match) {
            var invalids = document.querySelectorAll(':invalid'),
                result = false,
                len = invalids.length;

            if (len) {
                for (var i = 0; i < len; i++) {
                    if (elem === invalids[i]) {
                        result = true;
                        break;
                    }
                }
            }
            return result;
        }
    });

    const newRowHtml = '<tr class="form-row"><td><select class="addRemove form-control" name="addRemove"><option value="Add">Add</option><option value="Remove">Remove</option></select></td><td><input class="name form-control" type="text" placeholder="Name" name="name" required></td><td><input class="email form-control" type="email" placeholder="name@domain.com" name="email" required></td><td><input class="objectId form-control" type="text" placeholder="ObjectID" name="objectId" pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}" required></td><td><input class="category form-control" name="category" list="category-list" required></select></td><td><span class="glyphicon glyphicon-plus adder"></span></td></tr><div id="next-row"';

    var api_url = "https://prod-28.westcentralus.logic.azure.com:443/workflows/b2c3172f32f44fb0bd8c0aa4d088074f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7NG895qipyLxhThbWOS8eZDioXpv_0UEZbc1s4xJNrI";

    function disableItAdminsLink() {
        $('#getItAdmins').on("click", function (e) {
            e.preventDefault();
        });
    }

    function enableItAdminsLink() {
        $("#getItAdmins").click(function () {
            var tid = $('input.tenantId').val();
            if (tid.length != 36) {
                return;
            }

            $(this).html(spinner + $(this).html());

            $.ajax({
                type: 'GET',
                url: itAdminsApiUrl + tid,
                success: function (data) {
                    $('#getItAdmins').html($('#getItAdmins').html().replace(spinner, ""));
                    var adminList = data["itAdmins"];
                    var admins = adminList.split(";");
                    $('#admins').html("");
                    admins.forEach(function (admin) {
                        admin = admin.trim();
                        if (admin.length > 0) {
                            $("#admins").append('<li>' + admin + '</li>');
                        }
                    });
                }
            });
        })
    }

    function submitForm(e) {
        e.preventDefault();

        let table = document.querySelector("table")
        let rows = document.querySelectorAll('tr.form-row');
        var arr = [];

        function checkIfDone(arr) {
            if (arr.length == rows.length) {
                finalSend(arr);
            }
        }

        rows.forEach((row) => {
            console.log(row);
            let company = document.querySelector("input.company").value;
            let tenantId = document.querySelector("input.tenantId").value;

            let addRemove = row.querySelector("select.addRemove").value;
            let name = row.querySelector("input.name").value;
            let email = row.querySelector("input.email").value;
            let objectId = row.querySelector("input.objectId").value;
            let category = row.querySelector("input.category").value;

            microsoftTeams.getContext(function (context) {
                let body = {
                    company: company,
                    tenantId: tenantId,
                    userEmail: context["userPrincipalName"],

                    userOrTenant: "User",
                    ring: "1.5",

                    addRemove: addRemove,
                    name: name,
                    email: email,
                    objectId: objectId,
                    category: category,
                }
                arr.push(body);

                checkIfDone(arr);
            });
        
        });

        return false;
    }

    function finalSend(arr) {
        console.log("Sending the array");
        $('.alert-success').css('display', '');
        ajaxRequest('post', api_url, arr, printFormResults);
    }

    function printFormResults(e) {
        console.log(e);
    }

    function addRow() {
        var table = document.querySelector('table#form-table');

        let rows = table.querySelectorAll(".form-row");
        console.log("Currently " + rows.length + " rows");
        if (rows.length == 10) {
            $('#limitText').css('display', '');
            return;
        }

        let adders = table.querySelectorAll('span.adder');
        adders.forEach((adder) => {
            adder.remove();
        });

        var newRow = table.insertRow();
        newRow.classList.add("form-row");
        newRow.innerHTML = newRowHtml;
        //newRow.querySelector("input.company").value = lastCompany;

        // Add event listener to new plus button
        var plusButton = newRow.querySelector('span.adder');
        console.log(plusButton);
        plusButton.addEventListener('click', addRow);

        // Add excelPaste handlers to the pastable fields
        newRow.querySelector('input.name').addEventListener('paste', excelPaste);
        newRow.querySelector('input.email').addEventListener('paste', excelPaste);
        newRow.querySelector('input.objectId').addEventListener('paste', excelPaste);
        newRow.querySelector('input.category').addEventListener('paste', excelPaste);

        // Recalculate rows to see if the adder should be removed
        rows = table.querySelectorAll(".form-row");
        if (rows.length >= 10) {
            console.log(plusButton)
            plusButton.style.display = 'none';
            $('#limitText').css('display', '');
        }

        $('#tenantIdField').change(function () {
            unlockViewAdminsButtonIfValid();
        })

        $('input').change(function () {
            unlockButtonIfValid();
        })

        // TODO: This isn't working yet
        $('input.objectId').change(function () {
            console.log("objectId field changed");
            var oId = $(this).val();
            var tId = $('input.tenantId').val();
            console.log(oId, tId);
        })

    }

    function removeRow() {

    }

    function excelPaste(e) {
        // Prevent the default pasting event and stop bubbling
        e.preventDefault();
        e.stopPropagation();

        // Get the clipboard data
        let paste = (e.clipboardData || window.clipboardData).getData('text');

        // Do something with paste like remove non-UTF-8 characters
        let thisField = document.activeElement;

        let currentColIndex = thisField.parentElement.cellIndex;
        let currentRowIndex = thisField.parentElement.parentElement.rowIndex - 1;

        let rowObjects = thisField.parentElement.parentElement.parentElement.children;
        let totalRows = rowObjects.length;

        let pasteRows = paste.split("\n");
        if (pasteRows[pasteRows.length - 1].length == 0) {
            pasteRows = pasteRows.slice(0, pasteRows.length - 1);
        }

        if (pasteRows.length > (totalRows - currentRowIndex)) {
            let insufficiency = pasteRows.length - (totalRows - currentRowIndex);
            for (let i = 0; i < insufficiency; i++) {
                addRow();
            }
        }

        for (let rowInd in pasteRows) {
            rowInd = parseInt(rowInd);
            let row = pasteRows[rowInd];
            let pasteCols = row.split("\t");
            for (let colInd in pasteCols) {
                colInd = parseInt(colInd);
                let col = pasteCols[colInd];

                let thisCell = rowObjects[currentRowIndex + rowInd].children[currentColIndex + colInd];
                thisCell.children[0].value = col;
            }
        }

        unlockButtonIfValid();
    }

    function setup() {
        var adder = document.querySelector("span.adder");
        adder.addEventListener('click', addRow);

        var submitter = document.querySelector('#submitForm');
        submitter.addEventListener('click', submitForm);

        //document.querySelector('input.company').addEventListener('paste', excelPaste);
        document.querySelector('input.name.form-control').addEventListener('paste', excelPaste);
        document.querySelector('input.email.form-control').addEventListener('paste', excelPaste);
        document.querySelector('input.objectId').addEventListener('paste', excelPaste);
        document.querySelector('input.category').addEventListener('paste', excelPaste);

        $('input').change(function () {
            $(this)[0].setCustomValidity("");
        });

        $('input').change(function () {
            unlockButtonIfValid();
        });

        $('input.objectId').change(function () {
            console.log("objectId field changed");
            var oId = $(this).val();
            var tId = $('input.tenantId').val();
            console.log(oId, tId);
        })

        $('#tenantIdField').change(function () {
            if ($(this).val().length == 36) {
                //$("#getItAdmins").prop('disabled', false);
                enableItAdminsLink();
            } else {
                //$("#getItAdmins").prop('disabled', true);
                disableItAdminsLink();
            }
        });

        $('#submitForm').prop('disabled', true);
    }

    function unlockViewAdminsButtonIfValid() {
        var valid = true;
        if ($('#tenantIdField').is(':invalid')) {
            valid = false;
        }
        $('#getItAdmins').prop('disabled', valid);
    }

    function unlockButtonIfValid() {
        var something_is_invalid = false;
        if ($('input').is(':invalid')) {
            console.log($(this) + " is invalid");
            
            something_is_invalid = true;
        }

        //console.log("Invalid state: " + something_is_invalid);

        $('#submitForm').prop('disabled', something_is_invalid);
    }

    function checkUserVsTenantId() {
    }

    function ajaxRequest(method, url, params, callback) {
        var xmlhttp = new XMLHttpRequest();

        xmlhttp.onreadystatechange = function () {
            if (xmlhttp.readyState === 4 && xmlhttp.status === 200) {
                callback(xmlhttp.response);
            }
        };

        xmlhttp.open(method, url, true);
        console.log("Stringified: " + JSON.stringify(params));
        xmlhttp.setRequestHeader('Content-Type', 'application/json');
        xmlhttp.send(JSON.stringify(params));
    }

    function ready(fn) {
        if (typeof fn !== 'function') {
            return;
        }

        if (document.readyState === 'complete') {
            return fn();
        }

        document.addEventListener('DOMContentLoaded', fn, false);
    }

    ready(setup);

})();