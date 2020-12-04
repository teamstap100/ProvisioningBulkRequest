'use strict';

(function () {
    var itAdminsApiUrl = "../api/tenants/";

    var spinner = '<i class="fa fa-spinner fa-spin"></i>  ';

    function cleanEmail(email) {
        console.log("Cleaning email");
        console.log(email);

        // Deal with undefined email
        if (!email) {
            return email;
        }

        email = email.toLowerCase();
        console.log(email);
        email = email.replace("#ext#@microsoft.onmicrosoft.com", "");
        console.log(email);
        if (email.includes("@")) {
            return email;

        } else if (email.includes("_")) {
            console.log("Going the underscore route");
            var underscoreParts = email.split("_");
            var domain = underscoreParts.pop();
            var tenantString = domain.split(".")[0];

            if (underscoreParts.length > 1) {
                email = underscoreParts.join("_") + "@" + domain;
            } else {
                email = underscoreParts[0] + "@" + domain;
            }
        }
        return email;
    }

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
    
    const newRowHtml = '<tr class="form-row"><td><select class="addRemove form-control" name="addRemove"><option value="Add">Add</option><option value="Remove">Remove</option></select></td><td><input class="name form-control" type="text" placeholder="Name" name="name" required=""></td><td><input class="email form-control" type="email" placeholder="name@domain.com" name="email" required=""></td><td><input class="objectId form-control" type="text" placeholder="ObjectID" name="objectId" title="Please enter a valid ObjectID." required="" pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}"></td><td><input class="position form-control" type="text" placeholder="Position" name="position"></td><td><input class="department form-control" type="text" placeholder="Department" name="department"></td><td><span class="glyphicon glyphicon-plus adder"></span></td></tr>';

    // R3 Bulk Provisioning flow. This sends confirmation email, checks the OIDs, and sends it to the "immmediately create a PR" flow
    //const api_url = "https://prod-00.westcentralus.logic.azure.com:443/workflows/f2260dc702c04364b5e6ec2d7901db72/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TEg-D2oPUX_eOw8WEW_cUqWmqbZ9up31Xq6_9vfp3Zk";

    // Normal bulk provisioning flow. This writes the requests to the Provisioning Requests sheet like usual requests.
    const api_url = "https://prod-28.westcentralus.logic.azure.com:443/workflows/b2c3172f32f44fb0bd8c0aa4d088074f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7NG895qipyLxhThbWOS8eZDioXpv_0UEZbc1s4xJNrI";

    var maxUsers = 50;

    function disableItAdminsLink() {
        $('#getR3Users').on("click", function (e) {
            e.preventDefault();
        });
    }

    function enableItAdminsLink() {
        $("#getR3Users").click(function () {
            var tid = $('input.tenantId').val();
            
            if (tid.length != 36) {
                return;
            }

            console.log(tid);

            $(this).html(spinner + $(this).html());

            $.ajax({
                type: 'GET',
                url: itAdminsApiUrl + tid,
                success: function (data) {
                    $('#getR3Users').html($('#getR3Users').html().replace(spinner, ""));
                    data = JSON.parse(data);
                    if (data.r3_users.length ) {
                        var adminList = data.r3_users
                        $('#r3Users').html("");
                        adminList.forEach(function (admin) {
                            admin = admin.trim();
                            if (admin.length > 0) {
                                $("#r3Users").append('<li>' + admin + '</li>');
                                maxUsers -= 1;
                                console.log(maxUsers);
                            }
                        });
                    } else {
                        $('#r3Users').html("No R3 users yet for this tenant.");
                    }

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

            let validation = "End User";

            try {
                validation = $('#validationSelect option:selected').text();
            } catch (e) {
                validation = "End User";
            }

            let addRemove = row.querySelector("select.addRemove").value;
            let name = row.querySelector("input.name").value;
            let email = row.querySelector("input.email").value;
            let objectId = row.querySelector("input.objectId").value;
            let position = row.querySelector("input.position").value;
            let department = row.querySelector("input.department").value;

            microsoftTeams.getContext(function (context) {
                let body = {
                    company: company,
                    tenantId: tenantId,
                    userEmail: context["userPrincipalName"],

                    userOrTenant: "User",
                    ring: "3",

                    addRemove: addRemove,
                    name: name,
                    email: email,
                    objectId: objectId,
                    position: position,
                    department: department,
                    category: validation,
                }

                // Trying to send email to their UPN won't work if they're a guest
                try {
                    body.userEmail = cleanEmail(body.userEmail);
                } catch (e) {
                    console.log(e);
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
        //console.log("Currently " + rows.length + " rows");
        //console.log("Max users: " + maxUsers);
        //console.log(rows);

        let usersToAdd = 0;

        rows.forEach(function (row) {
            //console.log(row);
            let addRemoveVal = $(row).find('.addRemove').val();
            if (addRemoveVal == "Add") {
                usersToAdd++;
            } else {
                usersToAdd--;
            }
        })
        //console.log(usersToAdd + " / " + maxUsers)
        if (usersToAdd == maxUsers) {
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

        // Recalculate rows to see if the adder should be removed
        rows = table.querySelectorAll(".form-row");
        if (usersToAdd >= maxUsers) {
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
            try {
                console.log(rowInd);
                rowInd = parseInt(rowInd);
                let row = pasteRows[rowInd];
                let pasteCols = row.split("\t");
                for (let colInd in pasteCols) {
                    colInd = parseInt(colInd);
                    let col = pasteCols[colInd];

                    let thisCell = rowObjects[currentRowIndex + rowInd].children[currentColIndex + colInd];
                    thisCell.children[0].value = col;
                }
            } catch (e) {
                console.log("Too many rows");
            }
        }

        unlockButtonIfValid();
    }

    function setup() {
        microsoftTeams.initialize();

        var adder = document.querySelector("span.adder");
        adder.addEventListener('click', addRow);

        var submitter = document.querySelector('#submitForm');
        submitter.addEventListener('click', submitForm);

        //document.querySelector('input.company').addEventListener('paste', excelPaste);
        document.querySelector('input.name.form-control').addEventListener('paste', excelPaste);
        document.querySelector('input.email.form-control').addEventListener('paste', excelPaste);
        document.querySelector('input.objectId').addEventListener('paste', excelPaste);

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
            //console.log(oId, tId);
        })

        /*
        $('#tenantIdField').change(function () {
            if ($(this).val().length == 36) {
                //$("#getR3Users").prop('disabled', false);
                enableItAdminsLink();
            } else {
                //$("#getR3Users").prop('disabled', true);
                disableItAdminsLink();
            }
        });
        */

        $.ajax({
            url: "https://tap-validation-tab-admin-2.azurewebsites.net/api/50x50/count",
            type: "GET",
            dataType: 'json',
            success: function (data) {
                console.log(data.total);
                if (data.total < 2500) {
                    microsoftTeams.getContext(function (context) {
                        console.log("Getting context");

                        let userEmail = context['userPrincipalName'];

                        $.ajax({
                            url: "/api/tenants",
                            type: "POST",
                            data: { email: userEmail },
                            dataType: 'json',
                            success: function (data) {
                                data = JSON.parse(data);
                                console.log(data);
                                if (data) {
                                    $('input.company').val(data.name);
                                    $('input.tenantId').val(data.tid);
                                    console.log("R3 users max is: " + data.r3_users_max);
                                    if (data.r3_users_max) {
                                        maxUsers = data.r3_users_max;
                                    }
                                    enableItAdminsLink();
                                }

                            },

                        });
                    });


                    $.ajax({
                        url: "https://tap-validation-tab-admin-2.azurewebsites.net/api/validations",
                        type: "GET",
                        dataType: 'json',
                        success: function (data) {
                            if (data.length > 0) {
                                data.forEach(function (validation) {
                                    $('#validationSelect').append("<option>" + validation.name + "</option>");
                                })
                            }
                        },

                    });
                } else {
                    console.log("Program full");
                    $('#program-full').show();
                    $('#form').hide();
                    $('#submit-form').hide();
                }
            },

        });
    }

    function unlockViewAdminsButtonIfValid() {
        var valid = true;
        //if ($('#tenantIdField').is(':invalid')) {
        //    valid = false;
        //}
        $('#getR3Users').prop('disabled', valid);
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