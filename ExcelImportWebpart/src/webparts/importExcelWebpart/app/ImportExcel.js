var waiting;
var siteUrl = localStorage.getItem("url");
var requestDigest;
var totalCount = 0;
var countStudent = 0;
var countTeacher = 0;
var countClass = 0;
var location = "";
jQuery(document).ready(function () {

    var script2 = document.createElement('script');
    script2.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.7.7/xlsx.core.min.js';
    script2.type = 'text/javascript';
    document.getElementsByTagName('head')[0].appendChild(script2);

    var script3 = document.createElement('script');
    script3.src = 'https://cdnjs.cloudflare.com/ajax/libs/xls/0.7.4-a/xls.core.min.js';
    script3.type = 'text/javascript';
    document.getElementsByTagName('head')[0].appendChild(script3);
    bindLocationDetails();
    $('#viewfile').click(function () {
        if (validate()) {
            $('.Error').hide();
            getFormDgst();
        }
    });
    $("#btnClear").click(function () {
        $("#excelfile").val("");
        $("#ddtype1").val(0);
        $("#ddlLocation").val(0);
        $("#rowLocation").hide();
        $(".Error").hide();
    });
    $("#ddtype1").change(function () {
        if ($(this).val() == "Classes") {
            $("#rowLocation").show();
        } else {
            $("#rowLocation").hide();
            $("#ddlLocation").val(0);
            $("#spnLocationError").hide();
        }
    });
});
function validate() {
    var isValid = false;
    if ($("#ddtype1").val() != "0") {
        $("#spnListError").hide();
        isValid = true;
    }
    else {
        $("#spnListError").show();
        isValid = false;
    }
    if ($("#ddtype1").val() == "Classes") {
        if ($("#ddlLocation").val() != "0") {
            $("#spnLocationError").hide();
            isValid = true;
        } else {
            $("#spnLocationError").show();
            isValid = false;
        }
    }
    if ($("#excelfile").val() == 0) {
        $("#spnFileError").show();
        isValid = false;
    } else {
        $("#spnFileError").hide();
        isValid = true;
    }
    return isValid;
}
function getFormDgst() {
    var deferred = $.Deferred();
    $.ajax({
        url: siteUrl + "/_api/contextinfo",
        method: "POST",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function (data) {
            requestDigest = data.d.GetContextWebInformation.FormDigestValue;
            workOnIt();
            ExportToTable();
            deferred.resolve("true");
        },
        error: function (data, errorCode, errorMessage) {
            alert(errorMessage)
            deferred.reject("false");
            close();
        }
    });
    return deferred.promise();
}

function bindLocationDetails() {
    var apiPath = siteUrl + "/_api/lists/getbytitle('Location')/items";
    RestApiGet(apiPath).done(function (data) {
        var ddlLocation = $("#ddlLocation");
        if (data.length > 0) {
            ddlLocation.append("<option value='0'>Select Location</option>");
            data.forEach(element => {
                var htmlOption = "<option value=" + element.Title + ">" + element.Title + "</option>";
                ddlLocation.append(htmlOption);
            });
        }
    });
}
function RestApiGet(apiPath) {
    var deferred = $.Deferred();
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            var items; // Data will have user object  
            var results;
            if (data != null) {
                items = data.d;
                if (items != null) {
                    results = items.results;
                    deferred.resolve(results);
                }
            }
        },
        eror: function (data) {
            console.log("An error occurred. Please try again.");
            deferred.reject(0);
        }
    });
    return deferred.promise();
}
function ExportToTable() {
    location = $("#ddlLocation").val();
    var deferred = $.Deferred();
    try {
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;
        if (regex.test($("#excelfile").val().toLowerCase())) {
            var xlsxflag = false;
            if ($("#excelfile").val().toLowerCase().indexOf(".xlsx") > 0) {
                xlsxflag = true;
            }
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();
                reader.onload = function (e) {
                    var data = e.target.result;
                    if (xlsxflag) {
                        var workbook = XLSX.read(data, { type: 'binary' });
                    }
                    else {
                        var workbook = XLS.read(data, { type: 'binary' });
                    }
                    var sheet_name_list = workbook.SheetNames;
                    var cnt = 0;
                    sheet_name_list.forEach(function (y) {
                        if (xlsxflag) {
                            var exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]);
                        }
                        else {
                            var exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);
                        }
                        totalCount = exceljson.length;
                        console.log(totalCount + "totalCount");
                        if (exceljson.length > 0 && cnt == 0) {

                            //BindTable(exceljson, '#exceltable');
                            exceljson.forEach(function (excelRow) {
                                if (excelRow != null && Object.keys(excelRow).length > 0) {
                                    var ddlValue = $("#ddtype1").val();
                                    if (ddlValue != "") {
                                        try {
                                            switch (ddlValue) {
                                                case 'Students':
                                                    try {
                                                        if (excelRow["Stud No"].toString().trim() != "" && excelRow["Stud No"].toString().trim() != undefined)
                                                            getStudentsDetailsByStudentsID(excelRow["Stud No"].toString().trim(), excelRow);
                                                    } catch (Exception) {
                                                        countStudent++;
                                                    }
                                                    break;
                                                case 'Teachers':
                                                    try {
                                                        if (excelRow["Email"].toString().trim() != "" && excelRow["Email"].toString().trim() != undefined)
                                                            getTeachersDetailsByTeacherID(excelRow["Email"].toString().trim(), excelRow);
                                                    } catch (Exception) {
                                                        countTeacher++;
                                                    }
                                                    break;
                                                case 'Classes':
                                                    try {
                                                        if (excelRow["Class"].toString().trim() != "" && excelRow["Class"].toString().trim() != undefined)
                                                            getClassDetailsByClassID(excelRow["Class"].toString().trim(), excelRow);
                                                    } catch (Exception) {
                                                        countClass++;
                                                    }
                                                    break;
                                            }
                                        }
                                        catch (Exception) {

                                        }
                                    }
                                }
                                cnt++;
                            });
                        }
                    });
                    $('#exceltable').show();
                }
                if (xlsxflag) {
                    reader.readAsArrayBuffer($("#excelfile")[0].files[0]);
                }
                else {
                    reader.readAsBinaryString($("#excelfile")[0].files[0]);
                }
            }
            else {
                close();
                alert("Sorry! Your browser does not support HTML5!");
            }
        }
        else {
            alert("Please upload a valid Excel file!");
            close();
        }
        deferred.resolve("true");
    } catch (Exception) {
        deferred.reject("false");
        clsoe();
    }
    return deferred.promise();
}
function BindTable(jsondata, tableid) {
    var columns = BindTableHeader(jsondata, tableid);
    for (var i = 0; i < jsondata.length; i++) {
        var row$ = $('<tr/>');
        for (var colIndex = 0; colIndex < columns.length; colIndex++) {
            var cellValue = jsondata[i][columns[colIndex]];
            if (cellValue == null)
                cellValue = "";
            row$.append($('<td/>').html(cellValue));
        }
        $(tableid).append(row$);
    }
}
function BindTableHeader(jsondata, tableid) {
    var columnSet = [];
    var headerTr$ = $('<tr/>');
    for (var i = 0; i < jsondata.length; i++) {
        var rowHash = jsondata[i];
        for (var key in rowHash) {
            if (rowHash.hasOwnProperty(key)) {
                if ($.inArray(key, columnSet) == -1) {
                    columnSet.push(key);
                    headerTr$.append($('<th/>').html(key));
                }
            }
        }
    }
    $(tableid).append(headerTr$);
    return columnSet;
}
function getTeachersDetailsByTeacherID(EmployeeIDValue, excelRow) {
    var objHeaders = {
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose"
        },
        async: false,
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'include'
    }
    fetch(siteUrl + "/_api/web/lists/GetByTitle('Teachers')/items?$filter=Email eq '" + EmployeeIDValue + "'&$select=*&$orderby=ID", objHeaders)
        .then(function (response) {
            return response.json()
        })
        .then(function (json) {
            var results = json.d.results;
            if (results.length > 0) {
                for (i in results) {
                    updateTeacherDetailsListItem(results[i].ID, excelRow);
                }
            }
            else {
                createTeacherDetailsListItem(excelRow);
            }
        })
        .catch(function (ex) {
            close();
            console.log("error");
        });
}
function updateTeacherDetailsListItem(itemID, excelRow) {
    $.ajax
        ({
            url: siteUrl + "/_api/web/lists/GetByTitle('Teachers')/items(" + itemID + ")",
            type: "POST",
            data: JSON.stringify
                ({
                    __metadata:
                    {
                        type: "SP.Data.TeachersListItem"
                    },
                    Title: excelRow["Teacher Name"],
                    Phone: excelRow["Phone"],
                    Email: excelRow["Email"],
                    TeacherLevel: excelRow["Level"],
                    Location: excelRow["Location"],
                    Status: excelRow["Status"]
                }),
            headers:
            {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
            },
            async: false,
            success: function (data, status, xhr) {
                countTeacher++;
                if (totalCount == countTeacher) {
                    close();
                    $("#excelfile").val("");
                    $("#ddtype1").val(0);
                    $("#ddlLocation").val(0);
                    $(".Error").hide();
                    alert("Data successfully imported!");
                    totalCount = 0;
                    countTeacher = 0;
                }
            },
            error: function (xhr, status, error) {
                close();
                console.log(xhr);
            }
        });
}
function createTeacherDetailsListItem(excelRow) {
    $.ajax
        ({
            url: siteUrl + "/_api/web/lists/GetByTitle('Teachers')/items",
            type: "POST",
            async: false,
            data: JSON.stringify
                ({
                    __metadata:
                    {
                        type: "SP.Data.TeachersListItem"
                    },
                    Title: excelRow["Teacher Name"],
                    Phone: excelRow["Phone"],
                    Email: excelRow["Email"],
                    TeacherLevel: excelRow["Level"],
                    Location: excelRow["Location"],
                    Status: excelRow["Status"]
                }),
            headers:
            {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "X-HTTP-Method": "POST"
            },
            success: function (data, status, xhr) {
                countTeacher++;
                if (totalCount == countTeacher) {
                    close();
                    $("#excelfile").val("");
                    $("#ddtype1").val(0);
                    alert("Data successfully imported!");
                    $("#ddlLocation").val(0);
                    $(".Error").hide();
                    totalCount = 0;
                    countTeacher = 0;
                }
            },
            error: function (xhr, status, error) {
                close();
                console.log(xhr);
            }
        });
}
function getClassDetailsByClassID(EmployeeIDValue, excelRow) {
    var objHeaders = {
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose"
        },
        async: false,
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'include'
    }
    fetch(siteUrl + "/_api/web/lists/GetByTitle('Classes')/items?$filter=Title eq '" + EmployeeIDValue + "' and Location eq '" + location + "' &$select=*&$orderby=ID", objHeaders)
        .then(function (response) {
            return response.json()
        })
        .then(function (json) {
            var results = json.d.results;
            if (results.length > 0) {
                for (i in results) {
                    updateClassDetailsListItem(results[i].ID, excelRow);
                }
            }
            else {
                createClassDetailsListItem(excelRow);
            }
        })
        .catch(function (ex) {
            close();
            console.log("error");
        });
}
function updateClassDetailsListItem(itemID, excelRow) {
    $.ajax
        ({
            url: siteUrl + "/_api/web/lists/GetByTitle('Classes')/items(" + itemID + ")",
            type: "POST",
            data: JSON.stringify
                ({
                    __metadata:
                    {
                        type: "SP.Data.ClassesListItem"
                    },
                    Title: excelRow["Class"],
                    Room: excelRow["Room"],
                    Teacher: excelRow["Teacher"],
                    ClassLevel: excelRow["Class Level"],
                    Program: excelRow["Program"],
                    ProficiencyTest: excelRow["Proficiency Test"],
                    Count: excelRow["Count"],
                    RoomCapacity: excelRow["Room Capacity"],
                    TextBook: excelRow["Text Book"],
                    Location: location
                }),
            headers:
            {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
            },
            async: false,
            success: function (data, status, xhr) {
                countClass++;
                if (totalCount == countClass) {
                    close();
                    $("#excelfile").val("");
                    $("#ddtype1").val(0);
                    alert("Data successfully imported!");
                    $("#ddlLocation").val(0);
                    $(".Error").hide();
                    totalCount = 0;
                    countClass = 0;
                }
            },
            error: function (xhr, status, error) {
                close();
                console.log(xhr);
            }
        });
}
function createClassDetailsListItem(excelRow) {
    $.ajax
        ({
            url: siteUrl + "/_api/web/lists/GetByTitle('Classes')/items",
            type: "POST",
            data: JSON.stringify
                ({
                    __metadata:
                    {
                        type: "SP.Data.ClassesListItem"
                    },
                    Title: excelRow["Class"],
                    Room: excelRow["Room"],
                    Teacher: excelRow["Teacher"],
                    ClassLevel: excelRow["Class Level"],
                    Program: excelRow["Program"],
                    ProficiencyTest: excelRow["Proficiency Test"],
                    Count: excelRow["Count"],
                    RoomCapacity: excelRow["Room Capacity"],
                    TextBook: excelRow["Text Book"],
                    Location: location
                }),
            headers:
            {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "X-HTTP-Method": "POST"
            },
            success: function (data, status, xhr) {
                countClass++;
                if (totalCount == countClass) {
                    close();
                    $("#excelfile").val("");
                    $("#ddtype1").val(0);
                    alert("Data successfully imported!");
                    $("#ddlLocation").val(0);
                    $(".Error").hide();
                    totalCount = 0;
                    countClass = 0;
                }
            },
            error: function (xhr, status, error) {
                close();
                console.log(xhr);
            }
        });
}
function getStudentsDetailsByStudentsID(EmployeeIDValue, excelRow) {
    var deferred = $.Deferred();
    var objHeaders = {
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose"
        },
        async: false,
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'include'
    }
    fetch(siteUrl + "/_api/web/lists/GetByTitle('Students')/items?$filter=Student_x0020_No eq '" + EmployeeIDValue + "'&$select=*&$orderby=ID", objHeaders)
        .then(function (response) {
            deferred.resolve(true);
            return response.json()
        })
        .then(function (json) {
            var results = json.d.results;
            if (results.length > 0) {
                for (i in results) {
                    updateStudentsDetailsListItem(results[i].ID, excelRow);
                }
            }
            else {
                createStudentsDetailsListItem(excelRow);
            }
        })
        .catch(function (ex) {
            close();
            deferred.reject(0);
            console.log("error");
        });
    deferred.resolve(true);
    return deferred.promise()
}
function updateStudentsDetailsListItem(itemID, excelRow) {

    $.ajax
        ({
            url: siteUrl + "/_api/web/lists/GetByTitle('Students')/items(" + itemID + ")",
            type: "POST",
            data: JSON.stringify
                ({
                    __metadata:
                    {
                        type: "SP.Data.StudentsListItem"
                    },
                    Student_x0020_No: excelRow["Stud No"],
                    Title: excelRow["Student Name"],
                    Email: excelRow["Email"],
                    Start: excelRow["Start"],
                    End: excelRow["End"],
                    Student_x0020_Level: excelRow["Student Level"],
                    Principal: excelRow["Principal"],
                    Principal_x0020__x002d__x0020_Ro: excelRow["Principal - Room"],
                    Principal_x0020__x002d__x0020_Te: excelRow["Principal - Teacher"],
                    isActive: true
                }),
            headers:
            {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
            },
            async: false,
            success: function (data, status, xhr) {
                countStudent++;
                if (totalCount == countStudent) {
                    close();
                    $("#excelfile").val("");
                    $("#ddtype1").val(0);
                    alert("Data successfully imported!");
                    $("#ddlLocation").val(0);
                    $(".Error").hide();
                    totalCount = 0;
                    countStudent = 0;
                }
            },
            error: function (xhr, status, error) {
                close();
                console.log(xhr);
            }
        });
}
function createStudentsDetailsListItem(excelRow) {

    $.ajax
        ({
            url: siteUrl + "/_api/web/lists/GetByTitle('Students')/items",
            type: "POST",
            data: JSON.stringify
                ({
                    __metadata:
                    {
                        type: "SP.Data.StudentsListItem"
                    },
                    Student_x0020_No: excelRow["Stud No"],
                    Title: excelRow["Student Name"],
                    Email: excelRow["Email"],
                    Start: excelRow["Start"],
                    End: excelRow["End"],
                    Student_x0020_Level: excelRow["Student Level"],
                    Principal: excelRow["Principal"],
                    Principal_x0020__x002d__x0020_Ro: excelRow["Principal - Room"],
                    Principal_x0020__x002d__x0020_Te: excelRow["Principal - Teacher"]
                }),
            headers:
            {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "X-HTTP-Method": "POST"
            },
            success: function (data, status, xhr) {
                countStudent++;
                if (totalCount == countStudent) {
                    close();
                    $("#excelfile").val("");
                    $("#ddtype1").val(0);
                    alert("Data successfully imported!");
                    $("#ddlLocation").val(0);
                    $(".Error").hide();
                    totalCount = 0;
                    countStudent = 0;
                }
            },
            error: function (xhr, status, error) {
                close();
                console.log(xhr);
            }
        });
}
function workOnIt() {
    $("#Loader").show();
}
function close() {
    $("#Loader").hide();
}