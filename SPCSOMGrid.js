/*
https://gist.github.com/shyamjyothi
Author: Shyam Jyothi
Date: 27-Nov-2016
Version 1.0
Parameter Description  (* ~ Mandatory Parameters)
    listName*: Name of the List which needs to be queried
    camlQuery*: string - CamlQuery for querying the List. Just mention the <Where> conditions
    heading: string - Heading for the Grid
    columns*: array - Columns information
        HeaderText* : Column Header
        InternalName* : List Field Internal Value
        IsAnchor* : true/false, if the column is hyperlinked
        ColumnType*: Text,Multiuser,Date
    sorting: string - true/false
    verticalScroll: string - true/false
    horizontalScroll: string - true/false
    dateFormat: Date Format
    href: url (ID values is defaulted as QueryString)
    sortColumn : string Internal Name of the column for sorting - Doesnt support Date and User
    SPPaginate : true/false SharePoint Paged Querying for faster Loading
    showSearch : true/false shows search box, this useful when SPPaginate is false
    loadMessage : Custom Loading message
    rowLimit*: no of rows to be displayed
*/


(function ($) {
     $.fn.SPCSOMGrid = function (userOptions) {
        var Maindiv = $(this);

        //Declaring Private Variables
        $.fn.SPCSOMGrid.defaults =
        {
            listName: "",
            camlQuery: "",
            heading: "",
            columns: "",
            MainDiv: Maindiv,
            sorting: false,
            verticalScroll: false,
            horizontalScroll: false,
            clientContext: null,
            includeColumns: '',
            dataSrc: '',
            rowLimit: 10,
            sortColumn: '',
            SPItem: null,
            SPList: null,
            gridDataArray: [], AllDataArray: [],
            SPCamlQuery: "",
            waitDialog: null,
            tableHtml: '',
            queryStringColumn : "ID",
            href: '',
            dateFormat: "dd-MMM-yyyy",
            loadMessage: "Loading search results",
            pageNo: 1,
            position: "",
            currPostion: "",
            nextPos: "",
            prevPos: "",
            firstLoad: true,
            ListItemCount: 1,
            ListItemCounter: 0,
            showSearch: false,
            SPPaginate: true,
            overrideThrottling: false,
            displayRowCount: 10,
            qsParam: "ID",
            hostedApp: false,
            hostweburl: "",
            appweburl: "",
            scriptbase: ""
        };
        
        var obj = $.extend({}, $.fn.SPCSOMGrid.defaults, userOptions);

        return this.each(function () {
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function Render() {
                RenderTable();
            });
        });

        function RenderTable(){   
            obj.waitDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose(obj.loadMessage, '', 90, 450);          
            if(Validate()) {
                try {
                    RenderFirstTime();
                }
                catch(err) {
                    obj.waitDialog.close();
                    obj.MainDiv.empty();
                    var msgHtml = MessageHtml(err.message);
                    obj.MainDiv.append(msgHtml);
                }
                           
            } 
                          
        }

        function RenderFirstTime() {
            if(obj.ListItemCount > 0) {
                RenderHeading(true);
                QuerySP(false);
            }
            else {
                obj.waitDialog.close();
                obj.MainDiv.empty();
                var msgHtml = MessageHtml(NO_DATA());
                obj.MainDiv.append(msgHtml) 
            }
        }

        // Validates the Options
        function Validate() {
            var success = true;
            if(obj.camlQuery == "") { success = false; ShowError(NO_Caml()) }
            if(obj.MainDiv == null) { success = false; ShowError(NO_MAIN_DIV()) }
            if(obj.listName == "") { success = false;ShowError(NO_LIST()); }
            if(obj.columns.length == 0) { success = false; ShowError(NO_COLS()); }
            if(obj.rowLimit == 0) { success = false; ShowError(NO_ROW_LIMIT()) }
            if(obj.overrideThrottling == true && obj.SPPaginate == false) {sucess = false; ShowError(THROTTLE_EXP()); }
            if(obj.SPPaginate == true) { obj.displayRowCount = obj.rowLimit;}         
            return success;
        }      

        //Generates the Table Heading nased on the columns entered
        function RenderHeading(isFirst){
              obj.tableHtml = "";  
             if (obj.columns.length > 0) {
                 var headerHtml = "";
                 headerHtml = '<table id="tblData" class="stripe cell-border" cellspacing="0" width="100%">';
                 var includeString = '';
                 headerHtml = headerHtml + '<thead>';
                 headerHtml = headerHtml + '<tr>';
                //bind headers
                for (var i = 0; i < obj.columns.length; i++) {
                    var headerText = obj.columns[i].HeaderText;
                    var internalName = obj.columns[i].InternalName;
                    var columnType = obj.columns[i].ColumnType;
                    if (internalName.toLowerCase() != "id" && internalName.toLowerCase() != "contenttype" && columnType != "Image") {
                        includeString = includeString == '' ? internalName : includeString + ',' + internalName;
                    }
                    headerHtml = headerHtml + '<th>' + headerText + '</th>';                    
                }
                includeString = includeString + ',ID,ContentType';
                obj.includeColumns = "Include(" + includeString + ")";
                headerHtml = headerHtml + "</tr></thead>";
                obj.tableHtml = obj.tableHtml + headerHtml;
             }
        }

        //Queries the SharePoint List to fetch results
        //Sample CamlQuery
        ////<View><Query><Where><Geq><FieldRef Name='ID' /><Value Type='Counter'>1</Value></Geq></Where><OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy></Query><RowLimit>10</RowLimit></View>";
        function QuerySP(isPaged) {      

            var orderby = '';
            var query = '';
            if(obj.overrideThrottling == true) {
                orderby = "<OrderBy UseIndexForOrderBy='TRUE' Override='TRUE'>";
            }
            if(obj.sortColumn != '') {
                orderby = "<OrderBy><FieldRef Name='" + obj.sortColumn + "' Ascending='False'/></OrderBy>" ;
            }
            else orderby = "<OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy>" ;

            var camlStr = "<View><Query>" ;
            camlStr = camlStr + obj.camlQuery + orderby + "</Query>";
            camlStr = camlStr + "<RowLimit>" + obj.rowLimit + "</RowLimit>";
            camlStr = camlStr +  "</View>" ;
            
           //Sets the Client Context based on AppModel
            GetClientContext()
            if(obj.clientContext !== null){
                obj.SPList = obj.clientContext.get_web().get_lists().getByTitle(obj.listName);
                obj.SPCamlQuery = new SP.CamlQuery();            
                obj.SPCamlQuery.set_viewXml(camlStr);
                if(isPaged) {
                    position = new SP.ListItemCollectionPosition();
                    position.set_pagingInfo(obj.position);
                    obj.SPCamlQuery.set_listItemCollectionPosition(position);
                }            
                obj.spItems = obj.SPList.getItems(obj.SPCamlQuery);
                obj.clientContext.load(obj.spItems, obj.includeColumns);
                obj.clientContext.executeQueryAsync(Function.createDelegate(this, onQuerySuccess), Function.createDelegate(this, onQueryFailure));
            }
        }

        //Delegate for Success
        function onQuerySuccess(sender, args) {
            var listEnumerator = obj.spItems.getEnumerator();
            var item; obj.gridDataArray = [];
            while (listEnumerator.moveNext()) {
                item = listEnumerator.get_current();
                obj.gridDataArray.push(item);
            }
            obj.ListItemCounter = obj.ListItemCounter + obj.gridDataArray.length;
            if(obj.gridDataArray.length == 0) {
                if(obj.firstLoad == false) { //Was at Last Page
                    obj.pageNo = obj.pageNo - 1;
                    obj.waitDialog.close();
                }
                else {
                    ShowError(NO_DATA()) 
                }
            }
            else {
                BindData();
                obj.waitDialog.close();
            }
            
        }

        //Delegate for failure
        function onQueryFailure(sender, args) {
            obj.waitDialog.close();
            obj.MainDiv.empty();
            var msgHtml = MessageHtml(args.get_message());
            obj.MainDiv.append(msgHtml);
        }

        //Loops the data queried and writed the Table Elements
        function BindData() {
            var gridData = obj.gridDataArray;            
            if(gridData.length > 0) {
                var rowHtml = "<tbody>";            
                for (var i = 0; i < gridData.length ; i ++) {
                    obj.MainDiv.attr("SPItem",gridData[i]);
                    rowHtml = rowHtml + "<tr>";
                    for (var j = 0; j < obj.columns.length; j++) {
                        var Id = String(gridData[i].get_item(obj.queryStringColumn));
                        var internalName = obj.columns[j].InternalName;
                        var isAnchor= obj.columns[j].IsAnchor;
                        var columnType = obj.columns[j].ColumnType;
                        var aHref = obj.href;
                        
                        
                        if (columnType != "Image") {
                            var columnValue = "";
                            var item = null;
                            item = internalName.toLowerCase() == "contenttype" ? gridData[i].get_contentType() : gridData[i].get_item(internalName);
                            switch (columnType) {
                                case 'Lookup':
                                    columnValue = item != null ? String(item.get_lookupValue()) : "";
                                    break;
                                case 'User':
                                    columnValue = item != null ? String(item.get_lookupValue()) : "";
                                    break;
                                case 'MultiUser':
                                    columnValue = item != null ? GetLookupInfo(item) : "";
                                    break;
                                case 'Date':
                                    columnValue = item != null ? String(item.format("dd-MMM-yyyy")) : "";
                                    break;
                                case 'Computed':
                                    columnValue = item != null ? item.get_name() : "";
                                    break;
                                case 'File' :
                                    columnValue = item != null ? String(item) : "";
                                    break;
                                default:
                                    columnValue = item != null ? String(item) : "";
                            }
                            if(isAnchor){
                                var qsParam;
                                if(obj.qsParam === '') qsParam = "ID";
                                else qsParam = obj.qsParam;
                                aHref = aHref + "?" + qsParam + "=" + Id
                                rowHtml = rowHtml + "<td>"+ "<a href='"+ aHref + "'>" + columnValue + "</a></td>";
                            }
                            else if(columnType.toLowerCase() == 'file') {  
                                aHref = aHref + columnValue;
                                rowHtml = rowHtml + "<td>"+ "<a href='"+ aHref + "' target='_blank'>" + RemoveFileExtn(columnValue) + "</a></td>";
                            }
                            else {rowHtml = rowHtml + "<td>" + columnValue + "</td>";}
                        
                        }
                }
                    rowHtml = rowHtml + "</tr>";
                }
               
               obj.tableHtml = obj.tableHtml + rowHtml + "</tbody></table>";              
               obj.MainDiv.html(obj.tableHtml);               

               //Binding of jQueryDatatable
                $("#tblData").DataTable( {
                    "ordering": obj.sorting,
                    "info": false,
                    "scrollX": obj.horizontalScroll,
                    "scrolly" : obj.verticalScroll,
                    "pageLength": obj.displayRowCount,
                    "dom": '<"toolbar">frtip'
                });
                if(obj.heading != '') $("div.toolbar").html(obj.heading); 
                if(!obj.showSearch) $("#tblData_filter").hide();
                //Custom Pagination Logic
                if(obj.SPPaginate) SetPagination();

            }
            else {
                    if(SPPaginate) {
                        if(obj.pageNo ==1 && obj.firstLoad) {
                            obj.MainDiv.empty();
                            var msgHtml = "";
                            var msgHtml = MessageHtml(NO_DATA());
                            obj.MainDiv.append(msgHtml);
                        }
                    }
                    else {
                        obj.MainDiv.empty();
                        var msgHtml = "";
                        var msgHtml = MessageHtml(NO_DATA());
                        obj.MainDiv.append(msgHtml); 
                    }
            }
            
        }


        function SetPagination() {
            //Clearing the Pagingation of jQuery Tables
            $("#tblData_info").hide();
            $("#tblData_paginate").hide();            

            //add hidden cntrls                
            var collection = obj.gridDataArray;
            var start = (obj.pageNo * obj.rowLimit) - (obj.rowLimit - 1);

            var end; var disableNext = false; var pclassname; var nclassname
           
            if(collection.length < obj.rowLimit) {  //this is the Last Page
                if(obj.firstLoad == true) { // items are less than rowlimit
                    end = obj.rowLimit ;
                    disableNext = true;
                }
                else {
                    end = ((obj.pageNo - 1) * obj.rowLimit) + collection.length ;
                    disableNext = true;
                }
            }


            if(collection.length == obj.rowLimit) { //may or maynot have LastPage
                end = (obj.pageNo) * obj.rowLimit;
            }           
             
            var pageMsg = start + " - " + end;
            
            if(obj.pageNo == 1)  pclassname = "paginate_button previous disabled";
            else pclassname = "paginate_button previous"

            if(disableNext)  nclassname = "paginate_button next disabled";
            else nclassname = "paginate_button next "

           
            var sortColumn = obj.sortColumn == '' ? "ID" : obj.sortColumn
            var ncolumnValue = String(collection[collection.length - 1].get_item(sortColumn));
            var pcolumnValue = String(collection[0].get_item(sortColumn));
            var npageFirstRow = end + 1;
            var ppageFirstRow = start;
            
            obj.nextPos = "Paged=TRUE&p_ID=" + String(collection[collection.length - 1].get_item('ID')) + "&p_" + sortColumn + "=" + ncolumnValue;
            obj.prevPos = "Paged=TRUE&PagedPrev=TRUE&p_ID=" + String(collection[0].get_item('ID')) +  "&p_" + sortColumn + "=" + pcolumnValue;   

           //footerdiv
            var footerDiv = $('<div class="dataTables_paginate paging_simple_numbers" id="tblData_SPCSOM"><a tabindex="0" class=" ' + pclassname     +'" id="tblData_previous_spcsom" aria-controls="tblData" data-dt-idx="0" href="#">Previous</a><span><a tabindex="0" class="paginate_button current" aria-controls="tblData" data-dt-idx="1">' + pageMsg + '</a></span><a tabindex="0" class="' + nclassname + '" id="tblData_next_spcsom" aria-controls="tblData" data-dt-idx="7" href="#">Next</a></div>');
            $("#tblData_wrapper").append(footerDiv);

            $("#tblData_previous_spcsom").click(function () {
                if(obj.pageNo > 1) {
                    PreviousPage();
                }
            });

            $("#tblData_next_spcsom").click(function () {
                if(!disableNext) {
                    NextPage();
                }
                
            });
        }

        function NextPage() {
            var pageNo =  obj.pageNo            
            obj.position = obj.nextPos;
            obj.pageNo = parseInt(pageNo) + 1;
            obj.firstLoad = false;
            try {
                    obj.waitDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose(obj.loadMessage, '', 90, 450);   
                    RenderHeading(false);
                    QuerySP(true); 
                }
                catch(err) {
                    obj.waitDialog.close();
                    obj.MainDiv.empty();
                    var msgHtml = MessageHtml(err.message);
                    obj.MainDiv.append(msgHtml);
                }          
        }

        function PreviousPage() {
            var pageNo =  obj.pageNo;           
            obj.position = obj.prevPos;
            obj.pageNo = parseInt(pageNo) - 1;
            obj.firstLoad = false;

            try {
                    obj.waitDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose(obj.loadMessage, '', 90, 450);   
                    RenderHeading(false);
                    QuerySP(true); 
                }
                catch(err) {
                    obj.waitDialog.close();
                    obj.MainDiv.empty();
                    var msgHtml = MessageHtml(err.message);
                    obj.MainDiv.append(msgHtml);
                }            
        }

        function GetClientContext(){
            if(!obj.hostedApp) {
                obj.clientContext = SP.ClientContext.get_current();
            }
            else {
                if(obj.clientContext !== null) {
                    hostweburl = GetQueryStringParameter("SPHostUrl");
                    appweburl = GetQueryStringParameter("SPAppWebUrl");
                    hostweburl = decodeURIComponent(hostweburl);
                    appweburl = decodeURIComponent(appweburl);
                    scriptbase = hostweburl + "/_layouts/15/";  
                    $.getScript(scriptbase + "SP.Runtime.js", function () {
                            $.getScript(scriptbase + "SP.js", function () {
                                 $.getScript(scriptbase + "SP.RequestExecutor.js", GetConextFromApp); });
                        });
                    }
                }
                
        }

        function GetConextFromApp() {
            obj.clientContext = new SP.ClientContext(appweburl);
            var factory =
                new SP.ProxyWebRequestExecutorFactory(
                    appweburl
                );
            obj.clientContext.set_webRequestExecutorFactory(factory);
            appContextSite = new SP.AppContextSite( obj.clientContext, hostweburl);
            var web = appContextSite.get_web();
            context.load(web);
        }


        function GetLookupInfo(userInfo) {
            var Users = "";
            if (typeof userInfo === "undefined") {
                Users = "";
            }
            else {
                if (typeof userInfo.length != "undefined") {
                    for (var i = 0; i < userInfo.length; i++) {
                        Users = Users != "" ? Users + " ; " + String(userInfo[i].get_lookupValue()) : String(userInfo[i].get_lookupValue());
                    }
                }
                else {
                    Users = String(userInfo.get_lookupValue());
                }
            }
        }

        function RemoveFileExtn(item) {
            var fileName = String(item)
            fileName = fileName.substr(0, fileName.lastIndexOf('.')) || fileName;            
            return fileName;
        }


        function MessageHtml(msg) {
            var messageHtml = "<div style='";
            messageHtml = messageHtml + "border: 1px solid;";
            messageHtml = messageHtml + "margin: 10px 0px;";
            messageHtml = messageHtml + "padding:15px 10px 15px 50px;";
            messageHtml = messageHtml + "background-repeat: no-repeat;";
            messageHtml = messageHtml + "background-position: 10px center;";
            messageHtml = messageHtml + "color: #9F6000;";
            messageHtml = messageHtml + "background-color: #FEEFB3;";
            messageHtml = messageHtml + "'>";
            messageHtml = messageHtml + msg;
            messageHtml = messageHtml + "</div>";
            return messageHtml;
        } 

        function ShowError(msg) {
            obj.waitDialog.close();
            obj.MainDiv.empty();
            var msgHtml = MessageHtml(msg);
            obj.MainDiv.append(msgHtml) 
        }

         function GetQueryStringParameter(paramToRetrieve) {
            var params =
                document.URL.split("?")[1].split("&");
            var strParams = "";
            for (var i = 0; i < params.length; i = i + 1) {
              var singleParam = params[i].split("=");
              if (singleParam[0] == paramToRetrieve)
                return singleParam[1];
            }
          }

        //Constant values
        function NO_DATA() { return "No Data Returned."}
        function NO_Caml() { return "CamlQuery parameter is empty."}
        function NO_LIST() { return "ListName is empty."}
        function NO_MAIN_DIV() { return "MainDiv is not specified."}
        function NO_COLS() { return "Columns not specified."}
        function NO_ROW_LIMIT() { return "RowLimit not specified."}
        function THROTTLE_EXP() { return "If the List is throttled please set SPPaginate to TRUE."}
        function ROWLIMIT_EXP() { return "If the List is throttled please set CAMLQuery Rowlimit and  Display Row Count to be same."}

//Ends

     };
})(jQuery)