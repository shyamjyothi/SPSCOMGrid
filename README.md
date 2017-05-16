# SPSCOMGrid

SPCSOMGrid is Client Side Grid View Component which uses SharePoint CSOM to displaying List/Library data. The grid internally used jQuery Datatable(freeware) for giving a rich look and feel on the data displayed. This can easily be plugged in to any HTML/ASPX pages by writing very minimal lines of JS code. With the use of SPCSOMGrid we don't have write any HTML for the displaying the information/data, we just have mention a container and SPCSOMGrid shall write the HTMLs internally. The SPCSOMGrid would require input parameters like CAMLQuery, List Name, Columns, Row Limit etc. based on which the Data is displayed SPCSOMGrid can be easily plugged in to various components as this is purely based on Javascript.

SPCSOMGrid primarily uses a JS file SPCSOMGrid.js which needs to be referred along with the jQuery and jQuery Datatable files in the HTML, ASPX etc. With the reference of the JS file, we have to create the parameters for the SPCSOMGrid and invoke the method against the container where we need to display the data grid.

The Sample use of the grid is shown in the file SPCSOMGrid.html

### Parameters Description

  listName*	String	Name of the List/Document Library

  camlQuery*	String	CamlQuery for querying the List. Only the Where condition need to be entered.

  Heading*	String	Heading of the Grid

  columns*	Array	Columns information on the Grid

  HeaderText*	String	Title of the column

  InternalName*	String	Internal Name of the column in SP List

  IsAnchor*	Boolean	True shall render the field as a hyperlink

 ColumnType*	String	The type of the Column in SP List. Supports the fields of Type Text, Multiuser, Date, Lookup, File, User.

 rowLimit	Integer	No of rows to be queried at a time. The Grid shall default show 10 rows.

 SPPaginate	Boolean	SharePoint Pagination shall be used and data shall be loaded faster. This must set to True for Large Lists.

 dateFormat	String	Date Format. Default format is “dd-MMM-yyyy”

 href	String	URL (ID values is defaulted as Query String). In the case of Libraries, the relative path of the Document Library.

 Sorting	Boolean	True/False. The columns values can be sorted.

 horizontalScroll	Boolean	This is useful when large no of columns is displayed. Please specify a fixed width for the container in which the Table is displayed.

 showSearch	Boolean	The Top Search shall be displayed on the basis of this param.

 loadMessage	String	Custom Loading message when data is being queried from the Lists.

 overrideThrottling	Boolean	If the List throttled.

 displayRowCount	Integer	no of rows to be displayed in the Grid
 
 ### Exception Handling
 Exceptions shall be displayed on the UI if any error occurs while querying the data or if any invalid parameters are passed.

