//Drag & Drop File Upload services 
//create by Yaniv Arditi 10.2018
//http://yanivrdt.wordpress.com/

(function (ns)
{
    //Members
    this.Constants = {
        ODATA_ENDPOINT_SUFFIX: "/api/data/v9.0/",
        ANNOTATION_ENTITY_NAME: "annotation",
        CREATE_NOTE_ACTION_NAME: "dyn_uploadFile",
        NO_HTML5_SUPPORT: "Your browser does not support the required features for drag & drop :\\",
        BLOCKED_FILE_EXTENSIONS: "ade;adp;app;asa;ashx;asmx;asp;bas;bat;cdx;cer;chm;class;cmd;com;config;cpl;crt;csh;dll;exe;fxp;hlp;hta;htr;htw;ida;idc;idq;inf;ins;isp;its;jar;js;jse;ksh;lnk;mad;maf;mag;mam;maq;mar;mas;mat;mau;mav;maw;mda;mdb;mde;mdt;mdw;mdz;msc;msh;msh1;msh1xml;msh2;msh2xml;mshxml;msi;msp;mst;ops;pcd;pif;prf;prg;printer;pst;reg;rem;scf;scr;sct;shb;shs;shtm;shtml;soap;stm;tmp;url;vb;vbe;vbs;vsmacros;vss;vst;vsw;ws;wsc;wsf;wsh",
        MAX_FILE_SIZE_MB: 5,
        NOTIFCATION_TYPE:
            {
                INFO: "infoMsg",
                SUCCESS: "successMsg",
                WARNING: "warningMsg",
                ERROR: "errorMsg"
            },
        MSG_ERR_MISSING_TITLE: "Please add note title",
        MSG_GENERAL_FAILURE: "Something went wrong :( ",
        MSG_PROCESSING: "Working on it...",
        NOTE_CREATION_SUCCESS: "Note created successfully :)",
        NOTE_CREATION_FAILURE: "Something went wrong :( Note creation failed",
        MSG_WRN_INVALID_EXT: "Invalid file type, can not upload :\\",
        MSG_WRN_INVALID_SIZE: "File too big, can not upload :\\",
        MSG_ERR_NO_CONTEXT: "No context found, operation aborted",
        DRAG_AREA_TEXT : "or drop file here"
    };

    //indicate user fed tile
    var isTitleFedByUser = false;

    //current selected file object details 
    var currentFile = null;
    var currentFileBody = null;

    //extract Global Context object 
    getContext = function ()
    {
        var result = null;

        if (typeof GetGlobalContext != "undefined")
        {
            result = GetGlobalContext();
        }
        else
        {
            if (typeof Xrm != "undefined")
            {
                result = Xrm.Page.context;
            }
            else
            {
                throw Constants.ERR_NO_CONTEXT;
            }
        }

        return result;
    }

    //extract query string parameters into values array
    getExternalParams = function ()
    {
        var result = null;

        if (location.search != "")
        {
            paramsArr = new Array();
            paramsArr = location.search.substr(1).split("&");

            for (var i in paramsArr)
            {
                paramsArr[i] = paramsArr[i].replace(/\+/g, " ").split("=");
            }

            result = paramsArr;

            return result;
        }
    }

    //extract external parameter value by parameter name  
    getExternalParamValue = function (ParamsArray, paramName)
    {
        var result = null;

        for (var i in ParamsArray)
        {
            if (ParamsArray[i][0].toLowerCase() == paramName)
            {
                result = (ParamsArray[i][1]);
                break;
            }
            else if (ParamsArray[i][0].toLowerCase() == "data")
            {
                var dataValues = decodeURIComponent(ParamsArray[i][1]).split("&");
                for (var i in dataValues)
                {
                    dataValues[i] = dataValues[i].replace(/\+/g, " ").split("=");

                    for (var j in dataValues)
                    {
                        if (dataValues[j][0] == paramName)
                        {
                            result = dataValues[j][1];
                            break;
                        }
                    }
                }
            }
        }

        return result;
    }

    // Convert Array Buffer to Base 64 string
    arrayBufferToBase64 = function (buffer)
    {
        var binary = '';
        var bytes = new Uint8Array(buffer);
        var len = bytes.byteLength;
        for (var i = 0; i < len; i++)
        {
            binary += String.fromCharCode(bytes[i]);
        }

        return window.btoa(binary);
    }

    //display notifcation 
    notifyUser = function (message, type, isFadeOut)
    {
        $("#notificationArea").removeClass();
        $("#notificationArea").addClass(type);
        $("#notificationArea").text(message);
        $("#notificationArea").show();

        //fade out message if required         
        if (isFadeOut)
        {
            $("#notificationArea").fadeIn(300).delay(3000).fadeOut(800);
        }
    }

    //reset notifications area
    clearNotifcationArea = function ()
    {
        document.querySelector('#notificationArea').innerHTML = "";
    }

    //handle note creation success
    handleNoteCreationSuccess = function ()
    {
        notifyUser(Constants.NOTE_CREATION_SUCCESS, Constants.NOTIFCATION_TYPE.SUCCESS, true);
    }

    //handle note creation failure
    handleNoteCreationFailure = function (error)
    {
        notifyUser(Constants.MSG_GENERAL_FAILURE, Constants.NOTIFCATION_TYPE.ERROR, true);
    }

    //create record using Web API create action  
    createRecord = function (record, entityName, successHandler, failureHandler)
    {
    debugger
        //set OData end-point
        var ODataEndpoint = getContext().getClientUrl() +
            Constants.ODATA_ENDPOINT_SUFFIX +
            entityName + "s";

        //define request
        var req = new XMLHttpRequest();
        req.open("POST", ODataEndpoint, true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function ()
        {
            if (this.readyState == 4)
            {
                req.onreadystatechange = null;

                if (this.status == 204)
                {
                    //handle success
                    successHandler();
                }
                else
                {
                    //handle failure
                    var error = JSON.parse(this.response).error;
                    failureHandler(error);
                }
            }
        };

        //send request 
        req.send(window.JSON.stringify(record));
    }

    //create record using Web API via a custom action  
    createRecordViaAction = function (record, entityName, successHandler, failureHandler)
    {
        //set OData end-point
        var ODataEndpoint = getContext().getClientUrl() +
            Constants.ODATA_ENDPOINT_SUFFIX +
            entityName + "s";

        //define request
        var req = new XMLHttpRequest();
        req.open("POST", ODataEndpoint, true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function ()
        {
            if (this.readyState == 4)
            {
                req.onreadystatechange = null;

                if (this.status == 204)
                {
                    //handle success
                    successHandler();
                }
                else
                {
                    //handle failure
                    var error = JSON.parse(this.response).error;
                    failureHandler(error);
                }
            }
        };

        //send request 
        req.send(window.JSON.stringify(record));
    }

    //create attachment
    createAttachment = function (annotationData)
    {
        //notify processing 
        notifyUser(Constants.MSG_PROCESSING, Constants.NOTIFCATION_TYPE.INFO, false)

        //define annotation record along with related object association
        var annotation = {
            subject: annotationData.title,
            notetext: annotationData.description,
            documentbody: currentFileBody != null ? currentFileBody : null,
            filename: annotationData.file != null ? annotationData.file.name : null,
            mimetype: annotationData.file != null ? annotationData.file.type : null
        };

        //bind annotation to context entity if exists
        if ((annotationData.contextRecordId != null) &&
            (annotationData.contextEntityName != null))
        {
            annotationData.contextRecordId = annotationData.contextRecordId.replace("%7b", "").replace("%7d", "");

            var odataEntityPluralName = handleIrregularEntityName(annotationData.contextEntityName);
 
            annotation["objectid_" + annotationData.contextEntityName +"@odata.bind"] = "/" +
                odataEntityPluralName + "s(" + annotationData.contextRecordId + ")";            
        }

        //create annotation record
        createRecord(annotation, Constants.ANNOTATION_ENTITY_NAME,
            handleNoteCreationSuccess, handleNoteCreationFailure);
    }

    //clear form controls 
    clearFormControls = function ()
    {
        //clear current file 
        currentFileBody = null;

        document.querySelector('#title').value = "";
        document.querySelector('#description').value = "";
        document.querySelector("#fileselect").value = "";
        document.querySelector('#filedrag').innerHTML = Constants.DRAG_AREA_TEXT;
    }

    //returns true if attahment data is valid
    isValidAttachment = function(attachmentData)
    {
        var result = false;

        if ((attachmentData.title != null) && (attachmentData.title != ""))
        {
            result = true;
        }

        return result;
    }

    //handle note submit event
    handleSubmitEvent = function ()
    {
        try
        {
            //get containing record details if exists 
            var externalParams = getExternalParams();

            var attachmentData = {
                title: document.querySelector('#title').value,
                description: document.querySelector('#description').value,
                //file: document.querySelector("#fileselect").files[0],
                file: currentFile,
                contextRecordId: getExternalParamValue(externalParams, "id"),
                contextEntityName: getExternalParamValue(externalParams, "typename")
            };
                                            
            //if valid note, create attachment record
            if (isValidAttachment(attachmentData))
            {
                //create attachment 
                createAttachment(attachmentData);

                //clear form controls 
                clearFormControls();
            }
            else
            {
                notifyUser(Constants.MSG_ERR_MISSING_TITLE,
                    Constants.NOTIFCATION_TYPE.WARNING,
                    true);
                document.querySelector('#title').focus();
            }
        }
        catch (e)
        {
            alert(e);
        }
    }

    //match entity name to Web API entity name (account/accounts vs. opportunity/opprtunities)
    handleIrregularEntityName = function(entityName)
    {
        var result = entityName;

        switch(entityName) {
            case "opportunity":
                result = "opportunitie";
                break;
            default:
                break;
        }

        return result;
    }

    //returns true if file extension is valid
    isValidExtension = function (fileName)
    {
        var result = false;

        //extract extention
        var ext = fileName.slice((fileName.lastIndexOf(".") - 1 >>> 0) + 2);

        //test extention 
        if (!(new RegExp(ext)).test(Constants.BLOCKED_FILE_EXTENSIONS))
        {
            result = true;
        }

        return result;
    }

    //returns true if target file size (bytes converted to KB) is valid 
    isValidFileSize = function (size)
    {
        var result = false;

        if (size <= (parseInt(Constants.MAX_FILE_SIZE_MB) * 1024000))
        {
            result = true;
        }

        return result;
    }

    //set file name as note title if note title input empty   
    setNoteTitle = function(fileName)
    {
        if ((!isTitleFedByUser) || (document.querySelector('#title').value == ""))
        {
            isTitleFedByUser = false;
            document.querySelector('#title').value = fileName;
            clearNotifcationArea();
        }
    }

    //handle file selection event
    handleFileSelection = function (event)
    {       
        // cancel event and hover styling
        fileDragHover(event);

        //extract target file
        var inputFiles = event.target.files || event.dataTransfer.files;
        
        var reader = new FileReader();
        reader.onload = function ()
        {
            currentFileBody = arrayBufferToBase64(reader.result);
        };
        
        //test for valid file extensions 
        if (isValidExtension(inputFiles[0].name))
        {            
            //test for valid file size 
            if (isValidFileSize(inputFiles[0].size) == true)
            {
                //set file name as note title if empty   
                setNoteTitle(inputFiles[0].name);

                //set file name in drag area to denote selected file
                document.querySelector('#filedrag').innerHTML = inputFiles[0].name;

                //if file dragged, clear file input control
                if (event.dataTransfer.files)
                {
                    document.querySelector("#fileselect").value = "";
                }

                //set current file 
                currentFile = inputFiles[0];

                //parse file bytes
                reader.readAsArrayBuffer(inputFiles[0]);
            }
            else
            {
                clearFormControls();
                notifyUser(Constants.MSG_WRN_INVALID_SIZE, Constants.NOTIFCATION_TYPE.ERROR, true)
            }
        }
        else
        {
            clearFormControls();
            notifyUser(Constants.MSG_WRN_INVALID_EXT, Constants.NOTIFCATION_TYPE.ERROR, true)
        }
    };

    // file drag hover
    fileDragHover = function(e)
    {        
        e.stopPropagation();
        e.preventDefault();
        e.target.className = (e.type == "dragover" ? "hover" : "");
    }

    //handle user fed title event 
    handleTitleFeed = function()
    {
        isTitleFedByUser = true;
    }

    Init = function ()
    {
        //get reference to form elements 
        var submitbutton = document.querySelector('#createAttachment');
        var fileSelector = document.querySelector('#fileselect');
        var filedrag = document.querySelector('#filedrag');
        var title = document.querySelector('#title');

        //add event to indicate user fed title			
        title.addEventListener("keypress", handleTitleFeed, false);

        //add file drag & drop events
        filedrag.addEventListener('dragenter', fileDragHover, false);
        filedrag.addEventListener("dragover", fileDragHover, false);
        filedrag.addEventListener("dragleave", fileDragHover, false);
        filedrag.addEventListener("drop", handleFileSelection, false);

        filedrag.style.display = "block";

        //attach submit button event 
        submitbutton.addEventListener("click", handleSubmitEvent);

        //attach file selection event 
        fileSelector.addEventListener("change", handleFileSelection);
    }

    //test for browser support for required HTML5 features 
    if (window.File && window.FileList && window.FileReader)
    {
        Init();
    }
    else
    {
        notifyUser(Constants.NO_HTML5_SUPPORT, Constants.NOTIFCATION_TYPE.ERROR, false);
    }

})(window.DnDFileUpload = window.DnDFileUpload || {});

