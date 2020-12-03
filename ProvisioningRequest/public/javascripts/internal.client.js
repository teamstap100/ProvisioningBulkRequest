'use strict';

(function () {
    var spinner = '<i class="fa fa-spinner fa-spin"></i>  ';

    microsoftTeams.initialize();

    var submitterEmail = "?";

    microsoftTeams.getContext(function (context) {
        console.log("Got Teams context");
        submitterEmail = context["userPrincipalName"];
    });

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

    // register jQuery extension
    jQuery.extend(jQuery.expr[':'], {
        focusable: function (el, index, selector) {
            return $(el).is('a, button, :input, [tabindex]');
        }
    });

    $(document).on('keypress', 'input,select', function (e) {
        if (e.which == 13) {
            e.preventDefault();
            // Get all focusable elements on the page
            var $canfocus = $(':focusable');
            var index = $canfocus.index(this) + 1;
            if (index >= $canfocus.length) index = 0;

            var tableId = this.parentElement.parentElement.parentElement.parentElement.id;
            var skipLength;
            if (tableId == "user-table") {
                skipLength = 7;
            } else {
                skipLength = 5;
            }

            $canfocus.eq(index+skipLength).focus();
        }
    });

    //const userRowHtml = '<tr class="form-row"><td><input class="addRemove form-control" type="text" name="addRemove" pattern="(Add)|(Remove)"></td><td><input class="ring form-control" type="text" name="Ring" pattern="(1.5)|(3)"></td><td><input class="company form-control" type="text" placeholder="Company" name="company" autocomplete="off"></td><td><input class="tenantId form-control" type="text" placeholder="Tenant ID" name="tenantId" autocomplete="off" pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}"></td><td><input class="name form-control" type="text" placeholder="Name" name="name"></td><td><input class="email form-control" type="email" placeholder="name@domain.com" name="email"></td><td><input class="objectId form-control" type="text" placeholder="ObjectID" name="objectId" title="Please enter a valid ObjectID." pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}"></td><td><input class="category form-control" name="category" list="user-category-list"><datalist id="user-category-list"><option value="IT Admin"></option><option value="Room Account"></option><option value="Device Account"></option><option value="Super User"></option><option value="Elite User"></option></datalist></td><td><span class="glyphicon glyphicon-plus adder"></span></td></tr>';

    const userRowHtml = '<tr class="form-row"><td><input class="addRemove form-control" type="text" name="addRemove" pattern="(Add)|(Remove)" value="Add"></td><td><input class="ring form-control" type="text" name="Ring" pattern="(1.5)|(3)" value="1.5"></td><td><input class="tenantId form-control" type="text" placeholder="Tenant ID" name="tenantId" autocomplete="off" pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}"></td><td><input class="name form-control" type="text" placeholder="Name" name="name"></td><td><input class="email form-control" type="email" placeholder="name@domain.com" name="email"></td><td><input class="objectId form-control" type="text" placeholder="ObjectID" name="objectId" title="Please enter a valid ObjectID." pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}"></td><td><input class="category form-control" name="category" list="user-category-list"><datalist id="user-category-list"><option value="IT Admin"></option><option value="Room Account"></option><option value="Device Account"></option><option value="Super User"></option><option value="Elite User"></option></datalist></td><td><input class="position form-control" type="text" placeholder="Position" name="position"></td><td><input class="department form-control" type="text" placeholder="Department" name="department"></td><td><span class="glyphicon glyphicon-plus adder"></span></td></tr>';

    const tenantRowHtml = '<tr class="form-row"><td><input class="addRemove form-control" type="text" name="addRemove" pattern="(Add)|(Remove)"></td><td><input class="ring form-control" type="text" name="Ring" pattern="(1.5)|(3)"></td><td><input class="name form-control" type="text" placeholder="Name" name="name"></td><td><input class="domain form-control" type="text" placeholder="domain.onmicrosoft.com" name="domain"></td><td><input class="objectId form-control" type="text" placeholder="Tenant OID" name="objectId" title="Please enter a valid ObjectID." pattern="([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}"></td><td><input class="category form-control" name="category" list="tenant-category-list"><datalist id="tenant-category-list"><option value="Prod Tenant"></option><option value="Dev Tenant"></option><option value="Test Tenant"></option></datalist></td><td><span class="glyphicon glyphicon-plus adder"></span></td></tr>';

    var api_url = "https://prod-28.westcentralus.logic.azure.com:443/workflows/b2c3172f32f44fb0bd8c0aa4d088074f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7NG895qipyLxhThbWOS8eZDioXpv_0UEZbc1s4xJNrI";

    function submitForm(e) {
        e.preventDefault();

        $('#submitForm').html(spinner + " Submit");

        let userTable = document.querySelector("#user-table");
        let tenantTable = document.querySelector("#tenant-table");

        let rows = document.querySelectorAll('tr.form-row');

        let userRows = userTable.querySelectorAll('tr.form-row');
        let tenantRows = tenantTable.querySelectorAll('tr.form-row');

        console.log(userRows);
        console.log(tenantRows);
        var arr = [];

        var done = 0;

        function checkIfDone() {
            console.log(done + " / " + rows.length);
            if (done == rows.length) {
                finalSend(arr);
            }
        }

        userRows.forEach((row) => {
            console.log(row);

            let addRemove = row.querySelector("input.addRemove").value;
            let ring = row.querySelector("input.ring").value;
            //let company = row.querySelector("input.company").value;
            let tenantId = row.querySelector("input.tenantId").value;
            let name = row.querySelector("input.name").value;
            let email = row.querySelector("input.email").value;
            let objectId = row.querySelector("input.objectId").value;
            let category = row.querySelector("input.category").value;
            let position = row.querySelector("input.position").value;
            let department = row.querySelector("input.department").value;

            // Make sure the really important ones are filled in
            if ((objectId == "") || (addRemove == "") || (ring == "") || (tenantId == "") || (email == "")) {
                console.log("This one is missing something important");
                done += 1;
                console.log("Done incremented, user failure");
                checkIfDone();
            } else {
                let body = {
                    //company: company,
                    tenantId: tenantId,
                    userEmail: submitterEmail,

                    userOrTenant: "User",
                    ring: ring,

                    addRemove: addRemove,
                    name: name,
                    email: email,
                    objectId: objectId,
                    category: category,
                    position: position,
                    department: department,
                }

                console.log(body);

                arr.push(body);
                done += 1;
                    
                console.log("Done incremented, user success");
                checkIfDone();
            }
        });

        tenantRows.forEach((row) => {
            console.log(row);

            let addRemove = row.querySelector("input.addRemove").value;
            let ring = row.querySelector("input.ring").value;
            let name = row.querySelector("input.name").value;
            let domain = row.querySelector("input.domain").value;
            let objectId = row.querySelector("input.objectId").value;
            let category = row.querySelector("input.category").value;
            

            // Make sure the really important ones are filled in
            if ((objectId == "") || (addRemove == "") || (ring == "")) {
                console.log("This tenant is missing something important");
                done += 1;
                console.log("Done incremented, tenant failure");
                checkIfDone();
            } else {
                let body = {
                    company: "",
                    tenantId: "",
                    userEmail: submitterEmail,

                    userOrTenant: "Tenant",
                    ring: ring,

                    addRemove: addRemove,
                    name: name,
                    email: domain,
                    objectId: objectId,
                    category: category
                }
                arr.push(body);
                done += 1;
                console.log("Done incremented, tenant success");
                checkIfDone();
            }
        });

        return false;
    }

    function finalSend(arr) {
        $('.alert-success').css('display', '');
        $('#submitForm').html("Submit");
        window.scrollTo({ top: 0, behavior: 'smooth' });
        ajaxRequest('post', api_url, arr, printFormResults);
    }

    function printFormResults(e) {
        console.log(e);
    }

    function addRowUsers() {
        var table = document.querySelector('table#user-table');

        let adders = table.querySelectorAll('span.adder');
        adders.forEach((adder) => {
            adder.remove();
        });

        // Get the last row's "company" value so it can be copied
        let addRemoves = table.querySelectorAll('input.addRemove');
        let lastAddRemove = "";
        addRemoves.forEach((row) => {
            lastAddRemove = row.value;
        });

        let rings = table.querySelectorAll('input.ring');
        let lastRing = "";
        rings.forEach((row) => {
            lastRing = row.value;
        });

        var newRow = table.insertRow();
        newRow.classList.add("form-row");
        newRow.innerHTML = userRowHtml;
        //newRow.querySelector("input.company").value = lastCompany;

        // Add event listener to new plus button
        var plusButton = newRow.querySelector('span.adder');
        //console.log(plusButton);
        plusButton.addEventListener('click', addRowUsers);

        // Add excelPaste handlers to the pastable fields
        newRow.querySelector('input.addRemove').addEventListener('paste', excelPaste);
        newRow.querySelector('input.ring').addEventListener('paste', excelPaste);
        //newRow.querySelector('input.company').addEventListener('paste', excelPaste);
        newRow.querySelector('input.tenantId').addEventListener('paste', excelPaste);
        newRow.querySelector('input.name').addEventListener('paste', excelPaste);
        newRow.querySelector('input.email').addEventListener('paste', excelPaste);
        newRow.querySelector('input.objectId').addEventListener('paste', excelPaste);
        newRow.querySelector('input.category').addEventListener('paste', excelPaste);
        newRow.querySelector('input.position').addEventListener('paste', excelPaste);
        newRow.querySelector('input.department').addEventListener('paste', excelPaste);

        newRow.querySelector('input.addRemove').value = lastAddRemove;
        newRow.querySelector('input.ring').value = lastRing;

        $('input').change(function () {
            unlockButtonIfValid();
        });
    }

    function addRowTenants() {
        var table = document.querySelector('table#tenant-table');

        let adders = table.querySelectorAll('span.adder');
        adders.forEach((adder) => {
            adder.remove();
        });

        // Get the last row's "company" value so it can be copied
        let addRemoves = table.querySelectorAll('input.addRemove');
        let lastAddRemove = "";
        addRemoves.forEach((row) => {
            lastAddRemove = row.value;
        });

        let rings = table.querySelectorAll('input.ring');
        let lastRing = "";
        rings.forEach((row) => {
            lastRing = row.value;
        });

        var newRow = table.insertRow();
        newRow.classList.add("form-row");
        newRow.innerHTML = tenantRowHtml;
        //newRow.querySelector("input.company").value = lastCompany;

        // Add event listener to new plus button
        var plusButton = newRow.querySelector('span.adder');
        //console.log(plusButton);
        plusButton.addEventListener('click', addRowTenants);

        // Add excelPaste handlers to the pastable fields
        console.log(newRow);
        newRow.querySelector('input.addRemove').addEventListener('paste', excelPaste);
        newRow.querySelector('input.ring').addEventListener('paste', excelPaste);
        newRow.querySelector('input.name').addEventListener('paste', excelPaste);
        newRow.querySelector('input.domain').addEventListener('paste', excelPaste);
        newRow.querySelector('input.objectId').addEventListener('paste', excelPaste);
        newRow.querySelector('input.category').addEventListener('paste', excelPaste);

        newRow.querySelector('input.addRemove').value = lastAddRemove;
        newRow.querySelector('input.ring').value = lastRing;

        $('input').change(function () {
            unlockButtonIfValid();
        });
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

        let table = thisField.parentElement.parentElement.parentElement.parentElement;
        console.log(table.id);

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
                if (table.id == "user-table") {
                    addRowUsers();
                } else {
                    addRowTenants();
                }
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

        // This doesn't count as a change() event, so need to check all the fields now
        unlockButtonIfValid();
    }

    function setup() {
        var userAdder = document.querySelectorAll("span.adder")[0];
        userAdder.addEventListener('click', addRowUsers);

        var tenantAdder = document.querySelectorAll("span.adder")[1];
        tenantAdder.addEventListener('click', addRowTenants);

        var submitter = document.querySelector('#submitForm');
        submitter.addEventListener('click', submitForm);

        var userTable = document.querySelector('table#user-table');
        var tenantTable = document.querySelector('table#tenant-table');

        userTable.querySelector('input.addRemove').addEventListener('paste', excelPaste);
        userTable.querySelector('input.ring').addEventListener('paste', excelPaste);
        //userTable.querySelector('input.company').addEventListener('paste', excelPaste);
        userTable.querySelector('input.tenantId').addEventListener('paste', excelPaste);
        userTable.querySelector('input.name.form-control').addEventListener('paste', excelPaste);
        userTable.querySelector('input.email.form-control').addEventListener('paste', excelPaste);
        userTable.querySelector('input.objectId').addEventListener('paste', excelPaste);
        userTable.querySelector('input.category').addEventListener('paste', excelPaste);
        userTable.querySelector('input.position').addEventListener('paste', excelPaste);
        userTable.querySelector('input.department').addEventListener('paste', excelPaste);

        tenantTable.querySelector('input.addRemove').addEventListener('paste', excelPaste);
        tenantTable.querySelector('input.ring').addEventListener('paste', excelPaste);
        tenantTable.querySelector('input.name.form-control').addEventListener('paste', excelPaste);
        tenantTable.querySelector('input.domain.form-control').addEventListener('paste', excelPaste);
        tenantTable.querySelector('input.objectId.form-control').addEventListener('paste', excelPaste);
        tenantTable.querySelector('input.category.form-control').addEventListener('paste', excelPaste);

        $('input').change(function () {
            $(this)[0].setCustomValidity("");
        });

        $('input').change(function () {
            unlockButtonIfValid();
        });

        $('#submitForm').prop('disabled', true);

        /*
        $("input.objectId").change(function (e) {
            console.log($(this).val());
            var result = !new RegExp("([a-f]|[0-9]){8}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){4}-([a-f]|[0-9]){12}").test($(this).val());
            console.log(result);
                
        });
        */
    }

    function unlockButtonIfValid() {
        var something_is_invalid = false;
        if ($('input').is(':invalid')) {
            console.log($(this)[0] + " is invalid");
            
            something_is_invalid = true;
         }

        console.log("Invalid state: " + something_is_invalid);

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