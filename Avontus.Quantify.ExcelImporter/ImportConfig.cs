using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avontus.Quantify.ExcelImporter
{

    class ImportConfig
    {
        static private string xml = @"<?xml version=""1.0"" encoding=""UTF - 8""?>	
<imports>
    <import name = ""Global Lists"" identifier= ""Import: Global Lists"" required = ""List Name and Value/Name"" sheetorder = ""5"">
        <column name = ""List Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Value/Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Is Active"" type = ""bool"" required= ""false"" />        
        <column name = ""Description"" type = ""string"" maxlength = ""120"" required= ""false"" />
    </import>

	 <import name = ""Tax Rate"" identifier= ""Import: TaxRate"" required= ""Name"" sheetorder = ""10"" >>
        <column name = ""Name"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Is Active"" type = ""bool"" required = ""false"" />
        <column name = ""RefID"" type = ""string"" maxlength = ""255"" required = ""false"" />
        <column name = ""Tax Agency"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Description"" type = ""string"" maxlength = ""255"" required = ""false"" />
        <column name = ""Rate"" type = ""string"" maxlength = ""50"" required = ""false"" />
	</import>  

    <import name = ""Unit of Measure"" identifier = ""Import: Unit Of Measure"" 
            required = ""Name, Description, IsActive"" sheetorder = ""13"">
        <column name = ""Name""  type = ""string""  maxlength = ""100"" required = ""true"" />
        <column name = ""Description""  type = ""string""  maxlength = ""255""  required = ""true"" />
        <column name = ""Is Active"" type =""bool"" required =""true"" />
    </import>

    <import name = ""Product Category"" identifier= ""Import: Product Category"" required = ""Name"" sheetorder = ""15"">
        <column name = ""Parent Name"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Cost Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Revenue Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Description"" type = ""string"" maxlength = ""255"" required= ""false"" />
      </import>

    <import name = ""Product Catalog"" identifier= ""Import: Product Catalog"" required = ""Part Number"" sheetorder = ""20"">
        <column name = ""Part Number"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Description"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Weight"" type = ""real"" required= ""false"" />
        <column name = ""Model No"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Category"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""List"" type = ""decimal"" required= ""false"" />
        <column name = ""Cost"" type = ""decimal"" required= ""false"" />
        <column name = ""Service"" type = ""boolean"" required = ""false"" />
        <column name = ""Length"" type = ""decimal"" required= ""false"" />
        <column name = ""Width"" type = ""decimal"" required= ""false"" />
        <column name = ""Height"" type = ""decimal"" required= ""false"" />
        <column name = ""Custom 1"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""Custom 2"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""Custom 3"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""Manufacturer"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""Manufacturer's Part Number"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""ROP"" type = ""real"" required= ""false"" />
        <column name = ""ROQ"" type = ""real"" required= ""false"" />
    </import>

    <import name = ""Consumables Category"" identifier= ""Import: Consumables Category"" required = ""Name"" sheetorder = ""25"">
        <column name = ""Parent Name"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Cost Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Revenue Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Description"" type = ""string"" maxlength = ""255"" required= ""false"" />
    </import>

    <import name = ""Consumables Catalog"" identifier= ""Import: Consumables Catalog"" required = ""Part Number"" sheetorder = ""30"">
        <column name = ""Part Number"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Description"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Weight"" type = ""real"" required= ""false"" />
        <column name = ""Category"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Cost"" type = ""decimal"" required= ""false"" />
        <column name = ""Manufacturer"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""Manufacturer's Part Number"" type = ""string"" maxlength = ""50"" required=""false"" />
        <column name = ""ROP"" type = ""real"" required= ""false"" />
        <column name = ""ROQ"" type = ""real"" required= ""false"" />
    </import>
	
	<import name = ""Rent Adjustment"" identifier= ""Import: Rent Adjustment"" required = ""All but Job Site Name"" sheetorder = ""35"">
        <column name = ""Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Is Active"" type = ""bool"" required= ""true"" />  
        <column name = ""Description"" type = ""string"" maxlength = ""255"" required= ""false"" />
        <!-- column name = ""Job Site Name"" type = ""string"" maxlength = ""50"" required= ""false""  /  -->
        <column name = ""Rule Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
		<column name = ""Date Or Days"" type = ""string"" maxlength = ""50"" required= ""true"" />
		<column name = ""Start"" type = ""string"" maxlength = ""50"" required= ""true"" />
		<column name = ""End"" type = ""string"" maxlength = ""50"" required= ""true"" />
		<column name = ""Discount Or Markup"" type = ""string"" maxlength = ""50"" required= ""true"" />		
		<column name = ""Value"" type = ""real"" required= ""true"" />
    </import>

	<import name = ""Rate Profile"" identifier= ""Import: Rate Profile"" required = ""Name, Is Active and Part Number"" sheetorder = ""40"">
        <column name = ""Name"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Is Active"" type = ""bool"" required= ""true"" />  
        <column name = ""Description"" type = ""string"" maxlength = ""255"" required= ""false""   />
		<column name = ""Label For Header"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Rental Days"" type = ""integer""  required= ""true"" />
        <column name = ""Part Number"" type = ""string"" maxlength = ""50"" required= ""true"" />
        <column name = ""Rate"" type = ""decimal"" required= ""false"" />
        <column name = ""Rent Qty"" type = ""real"" required= ""false"" />
        <column name = ""Job Cost"" type = ""real"" required= ""false"" />
        <column name = ""Delivery Charge"" type = ""real"" required= ""false"" />
        <column name = ""Return Credit"" type = ""real"" required= ""false"" />
        <column name = ""Sell Price"" type = ""real"" required= ""false"" />
        <column name = ""Replacement Cost"" type = ""real"" required= ""false"" />
		<column name = ""Min.Days"" type = ""integer"" required= ""false"" />
        <column name = ""Rent Adjustment"" type = ""string"" maxlength = ""50"" required= ""false"" />		
     </import>

    <import name = ""Customer"" identifier= ""Import: Customer"" required = ""Name and Is Active"" sheetorder = ""45"">
        <column name = ""Name"" type = ""string"" maxlength = ""100"" required= ""true"" />
        <column name = ""Is Active"" type = ""bool"" required= ""true"" />        
        <column name = ""Number"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Accounting ID"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Email"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Web Site"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Business Phone"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Fax#"" type = ""string"" maxlength= ""50"" required= ""false"" />
        <column name = ""Business Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
    </import>

    <import name = ""Customer Contact"" identifier= ""Import: CustomerContact"" required= ""Customer Name, First, Last, Username, and Password"" sheetorder = ""50"">>
        <column name = ""Customer Name"" type = ""string"" maxlength = ""100"" required = ""true"" />
        <column name = ""First"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Last"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Title"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Email"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Phone"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Fax"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Cell"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Notes"" type = ""string"" maxlength = ""256"" required = ""false"" />
        <column name = ""Is Active"" type = ""bool"" required= ""true"" />
        <column name = ""User Name"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Password"" type = ""string"" maxlength = ""50"" required = ""true"" />
    </import>

    <import name = ""Order"" identifier= ""Import: Order"" required= ""Customer Name, Order Number, "" sheetorder = ""55"">>
        <column name = ""Customer Name"" type = ""string"" maxlength = ""100"" required = ""true"" />
        <column name = ""Order Number"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Work Order Number"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Order Is Active"" type = ""bool"" required= ""true"" />
        <column name = ""Order Description"" type = ""string"" maxlength = ""255"" required = ""false"" />
        <column name = ""Order Approver Username"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Order Reviewer Username"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Order Amount"" type = ""decimal"" required = ""false"" />
        <column name = ""Work Order Is Active"" type = ""bool"" required= ""false"" />
        <column name = ""Work Order Description"" type = ""string"" maxlength = ""255"" required = ""false"" />
        <column name = ""Work Order Approver Username"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Work Order Reviewer Username"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Work Order Amount"" type = ""decimal"" required = ""false"" />
    </import>

    <import name = ""Vendor"" identifier= ""Import: Vendor"" required = ""Name and Is Active"" sheetorder = ""60"">
        <column name = ""Name"" type = ""string"" maxlength = ""100"" required= ""true"" />
        <column name = ""Is Active"" type = ""bool"" required= ""true"" />        
        <column name = ""Number"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Accounting ID"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Email"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Web Site"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Business Phone"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Fax#"" type = ""string"" maxlength= ""50"" required= ""false"" />
        <column name = ""Business Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
    </import>

    <import name = ""Vendor Contact"" identifier= ""Import: VendorContact"" required= ""Vendor Name, First, Last, Username, and Password""
	  sheetorder = ""65"" >>
        <column name = ""Vendor Name"" type = ""string"" maxlength = ""100"" required = ""true"" />
        <column name = ""First"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Last"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Title"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Email"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Phone"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Fax"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Cell"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Notes"" type = ""string"" maxlength = ""256"" required = ""false"" />
        <column name = ""User Name"" type = ""string"" maxlength = ""50"" required = ""true"" />
        <column name = ""Password"" type = ""string"" maxlength = ""50"" required = ""true"" />
    </import>

	<import name = ""Job Site"" identifier =""Import: JobSite"" required =""Parent Location, Cusstomer Name, Name, Is Active, and Is Billable(should be false)"" sheetorder = ""70"">
        <column name = ""Parent Location"" type =""string"" maxlength =""100"" required =""true"" />
        <column name = ""Customer Name"" type =""string"" maxlength =""100"" required =""true"" />
        <column name = ""Name"" type =""string"" maxlength =""100"" required =""true"" />
        <column name = ""Number"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Accounting ID"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Is Active"" type =""bool"" required =""true"" />
        <column name = ""Is Billable"" type =""bool"" required =""true"" />        
        <column name = ""Description"" type =""string"" maxlength =""255"" required =""false"" />
        <column name = ""Track Scaffolds"" type =""bool"" required =""true"" />
        <column name = ""Track Activities"" type =""bool"" required =""true"" />
        <column name = ""Requests"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Activity List1 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Activity List1 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Activity List2 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Activity List2 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Activity List3 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Activity List3 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Activity Text Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Activity Text Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 1 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 1 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 2 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 2 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 3 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 3 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 4 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 4 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 5 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 5 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 6 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 6 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold List 7 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 7 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Date 1 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Date 1 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Date 2 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Date 2 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Text 1 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Text 1 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Text 2 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Text 2 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Text 3 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Text 3 Required"" type =""bool"" required =""false"" />
        <column name = ""Scaffold Yes/No 1 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Yes/No 2 Label"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Business Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Business Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Billing Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Street1"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Street2"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping City"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping State"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Zip/Postal Code"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping County/Parish"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Country"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Latitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Longitude"" type = ""string"" maxlength = ""50"" required= ""false"" />
		<column name = ""Shipping Elevation"" type = ""string"" maxlength = ""50"" required= ""false"" />
        <column name = ""Job Text1"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Job Text2"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Job YesNo1"" type = ""bool"" required = ""false"" />
        <column name = ""Job YesNo2"" type = ""bool"" required = ""false"" />
        <column name = ""Job List1"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Job List2"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Billing Method"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""First Invoice Date"" type = ""string"" maxlength = ""50"" required = ""false"" />        
        <column name = ""FATA Minimum Days Rent"" type = ""string"" maxlength = ""50"" required = ""false"" />        
        <column name = ""FATA Allow Override"" type = ""bool"" required = ""false"" />
        <column name = ""FATA Include Additional Charges"" type = ""bool"" required = ""false"" />
        <column name = ""Daily / Monthly"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Invoice Every Number of Days"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Invoice Monthly Cycle Type"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Invoice Monthly Day of Month"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Advance Issue Credits"" type = ""bool"" required = ""false"" />
        <column name = ""Advance Minimum Days Rent"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Advance Invoice Cycle Days"" type = ""bool"" required = ""false"" />
        <column name = ""Advance Cycle Common Date"" type = ""bool"" required = ""false"" />
        <column name = ""Advance Cycle Start Date"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Rate Profile"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Purchase Order"" type= ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Work Order"" type= ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Default New Shipments Billable"" type = ""bool"" required = ""false"" />
        <column name = ""Hide Zero Charges"" type = ""bool"" required = ""false"" />
        <column name = ""Tax Rate 1"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Tax Rate 2"" type = ""string"" maxlength = ""50"" required = ""false"" />
        <column name = ""Tax Consumables and Product Sales"" type = ""bool"" required = ""false"" />
        <column name = ""Tax Rent"" type = ""bool"" required = ""false"" />
        <column name = ""Tax Service Ticket Damage"" type = ""bool"" required = ""false"" />
        <column name = ""Tax Delivery &amp; Return"" type = ""bool"" required = ""false"" />
    </import>

    <import name = ""Request"" identifier = ""Import: Request"" 
            required = ""Job Site Name, Request Number, Notes, Status, Requested By Username, Created By User Name"" 
            sheetorder = ""75"">
        <column name = ""Job Site Name"" type =""string"" maxlength=""100"" required=""true"" />
        <column name = ""Request Number"" type =""string"" maxlength=""50"" required=""true"" />
        <column name = ""Notes"" type =""string"" maxlength=""255"" required=""true"" />
        <column name = ""Order"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Status"" type =""string"" maxlength=""50"" required=""true"" />
        <column name = ""Assigned To Username"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Request Date"" type =""date"" required=""true"" />
        <column name = ""Date Needed"" type =""date"" required=""true"" />
        <column name = ""Request List 1"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Request List 2"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Request List 3"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Request Text 1"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Request Text 2"" type =""string"" maxlength=""50"" required=""false"" />
        <column name = ""Request YesNo 1"" type =""bool"" required=""false"" />
        <column name = ""Request YesNo 2"" type =""bool"" required=""false"" />
    </import>

    <import name = ""Scaffold Job List"" identifier=""Import: Scaffold Job List"" required=""Job Site Name, List and Value"" 
            sheetorder = ""80"">
        <column name = ""Job Site Name"" type = ""string"" maxlength=""100"" required=""true"" />
        <column name = ""List"" type = ""string""  maxlength=""50"" required=""true"" />
        <column name = ""Value"" type = ""string"" maxLength=""50"" required=""true"" />
        <column name = ""Description"" type = ""string"" maxLength=""255"" required=""false"" />
    </import>

    <import name = ""Scaffold Tag"" identifier = ""Import: Scaffold Tag"" required = ""Job Site Name, Tag, and Status"" sheetorder = ""85"">
        <column name = ""Job Site Name""  type = ""string""  maxlength = ""100"" required = ""true"" />
        <column name = ""Purchase Order""  type = ""string""  maxlength = ""50""  required = ""false"" />
        <column name = ""Work Order""  type = ""string""  maxlength = ""50""  required = ""false"" />
        <column name = ""Status""  type = ""string""  maxlength = ""20""  required = ""true"" />
        <column name = ""Tag""  type = ""string""  maxlength = ""50""  required = ""true"" />
        <column name = ""Priority""  type = ""string""  maxlength = ""50""  required = ""false"" />
        <column name = ""Project""  type = ""string""  maxlength = ""50""  required = ""false"" />
        <column name = ""Step""  type = ""string""  maxlength = ""50""  required = ""false"" />
        <column name = ""Requested By"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Representative"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Requestor"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Foreman"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Builder"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Dismantler"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Inspector"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Planned Load Date""  type = ""date""  required = ""false"" />
        <column name = ""Planned Build Date""  type = ""date""  required = ""false"" />
        <column name = ""Planned Dismantle Date""  type = ""date""  required = ""false"" />
        <column name = ""Actual Load Date""  type = ""date""  required = ""false"" />
        <column name = ""Actual Build Date""  type = ""date""  required = ""false"" />
        <column name = ""Actual Dismantle Date""  type = ""date""  required = ""false"" />
        <column name = ""Notes""  type = ""date""  required = ""false"" />
        <column name = ""Location Notes""  type = ""date""  required = ""false"" />
        <column name = ""Length"" type = ""real"" required = ""false"" />
        <column name = ""Width"" type = ""real"" required = ""false"" />
        <column name = ""Height"" type = ""real"" required = ""false"" />
        <column name = ""No Of Legs"" type = ""real"" required = ""false"" />
        <column name = ""No of Decks"" type = ""real"" required = ""false"" />
        <column name = ""Base Elevation"" type = ""real"" required = ""false"" />
  		<column name = ""Scaffold List 1"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 2"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 3"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 4"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 5"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 6"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold List 7"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Date 1"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Date 2"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Text 1"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Text 2"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Text 3"" type =""string"" maxlength =""50"" required =""false"" />
        <column name = ""Scaffold Yes/No 1"" type =""boolean"" required =""false"" />
        <column name = ""Scaffold Yes/No 2"" type =""boolean"" required =""false"" />
    </import>
	
    <import name = ""Scaffold Tag Activity"" identifier = ""Import: Scaffold Tag Activity"" 
            required = ""Job Site Name, Tag, Activity Type, and Sumamry"" sheetorder = ""90"">
        <column name = ""Job Site Name""  type = ""string""  maxlength = ""100"" required = ""true"" />
        <column name = ""Tag""  type = ""string""  maxlength = ""50""  required = ""true"" />
        <column name = ""Scaffold Tag Activity Type"" type = ""string"" maxlength = ""50"" required = ""true""/>
        <column name = ""Summary"" type = ""string"" maxlength = ""50"" required = ""true""/>
        <column name = ""Detail"" type = ""string"" maxlength = ""255"" required = ""false""/>        
        <column name = ""Shift"" type = ""string"" required = ""false""/>
        <column name = ""Request Date"" type = ""date"" required = ""false""/>
        <column name = ""Planned Date"" type = ""date"" required = ""false""/>
        <column name = ""Actual Date"" type = ""date"" required = ""false""/>
        <column name = ""Planned Load Date"" type = ""date"" required = ""false""/>
        <column name = ""Actual Load Date"" type = ""date"" required = ""false""/>   
        <column name = ""Scaffold Activity List1"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Scaffold Activity List2"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Scaffold Activity List3"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Scaffold Activity Text"" type = ""string"" maxlength = ""50"" required = ""false""/>
        <column name = ""Planned Hours"" type = ""real"" required = ""false"" />
        <column name = ""Planned Cost"" type = ""decimal"" required = ""false"" />
        <column name = ""Actual Hours"" type = ""real"" required = ""false"" />
        <column name = ""Actual Cost"" type = ""decimal"" required = ""false"" />    
    </import>


    <import name = ""Unit Hour Rate Profile"" identifier = ""Import: Unit Hour Rate Profile"" 
            required = ""Job Site Name,	Unit Hour Rate Profile, IsActive"" sheetorder = ""100"">
        <column name = ""Job Site Name""  type = ""string""  maxlength = ""100"" required = ""true"" />
        <column name = ""Unit Hour Rate Profile""  type = ""string""  maxlength = ""50""  required = ""true"" />
        <column name = ""Is Active"" type =""bool"" required =""true"" />
        <column name = ""Unit of Measure"" type = ""string"" maxlength = ""50"" required = ""true""/>
        <column name = ""Unit Factor"" type = ""real"" required= ""false"" />
        <column name = ""Split"" type = ""real"" required= ""false"" />
        <column name = ""Price"" type = ""decimal"" required= ""false"" />
    </import>
    <import name = ""Multiplier"" identifier = ""Import: Multiplier"" 
            required = ""Unit Hour Rate Name, Name, IsActive"" sheetorder = ""110"">
        <column name = ""Job Site Name""  type = ""string""  maxlength = ""100"" required = ""true"" />
        <column name = ""Name""  type = ""string""  maxlength = ""50""  required = ""true"" />
        <column name = ""Value"" type = ""real"" required= ""true"" />        
        <column name = ""Is Active"" type =""bool"" required =""true"" />
    </import>

</imports>";

        public string RequiredColumns;
        public enum ImportColunnType
        {
            Unknown = 0,
            String = 1,
            Integer = 2,
            Real = 3,
            Date = 4,
            Boolean = 5,
            Decimal = 6
        }

        public enum ImportSheetType
        {
            Unknown = 0,
            Normalized = 1,
            Balances = 2,
            BranchOrLaydownBalance = 3,
            JobSiteBalance = 4,
            ScaffoldTagBalance = 5,
            NonStandard1 = 11
        }
        public short SheetOrder;

        public class ImportColumn
        {
            public string ColumnHeading;
            public ImportColunnType ColumnType;
            public int? MaximumLength;
            public bool IsRequired;
            public int Column;

            public ImportColumn(string columnHeading, string columnType, string maximumLength, string isRequired, int column = 0)
            {
                ColumnHeading = columnHeading;
                switch (columnType)
                {
                    case "string":
                        ColumnType = ImportColunnType.String;
                        break;
                    case "integer":
                        ColumnType = ImportColunnType.Integer;
                        break;
                    case "real":
                        ColumnType = ImportColunnType.Real;
                        break;
                    case "decimal":
                        ColumnType = ImportColunnType.Decimal;
                        break;
                    case "date":
                        ColumnType = ImportColunnType.Date;
                        break;
                    case "bool":
                        ColumnType = ImportColunnType.Boolean;
                        break;
                    default:
                        ColumnType = ImportColunnType.Unknown;
                        break;
                }
                MaximumLength = (maximumLength == "") ? null : (int?)Int32.Parse(maximumLength);
                IsRequired = isRequired == "true" ? true : false;
                Column = column;
            }

        }

        public string Name;
        //public ExcelImportType ImportType; 
        public string ImportIdentifier;
        public List<ImportColumn> ImportColumns = new List<ImportColumn>();

        static public List<ImportConfig> GetAllImportConfigs()
        {
            List<ImportConfig> list = new List<ImportConfig>();

            int columnNumber = 0;
            var doc = XDocument.Parse(xml);
            var imports = doc.Root.Elements();
            foreach (var import in imports)
            {
                columnNumber = 0;

                ImportConfig config = new ImportConfig();
                config.Name = import.Attribute("name").Value;
                config.RequiredColumns = import.Attribute("required").Value;
                config.SheetOrder = Int16.Parse(import.Attribute("sheetorder").Value.ToString());
                foreach (var e in import.Elements())
                {
                    string maxStr = e.Attribute("maxlength") == null ? "" : e.Attribute("maxlength").Value;
                    config.ImportIdentifier = import.Attribute("identifier").Value;

                    ImportColumn column = new ImportColumn(e.Attribute("name").Value,
                        e.Attribute("type").Value, maxStr,
                        e.Attribute("required").Value, columnNumber++);



                    config.ImportColumns.Add(column);
                }
                list.Add(config);
            }
            return list;
        }

        static public ImportConfig GetImportConfig(string targetName)
        {
            ImportConfig config = new ImportConfig();

            int columnNumber = 1;
            bool configLoaded = false;
            var doc = XDocument.Parse(xml);
            var imports = doc.Root.Elements();
            foreach (var import in imports)
            {
                columnNumber = 1;
                string importIdentifier = import.Attribute("identifier").Value;
                if (ImporterData.NormStr(importIdentifier) != targetName)
                {
                    continue;
                }
                configLoaded = true;
                config.RequiredColumns = import.Attribute("required").Value;

                foreach (var e in import.Elements())
                {
                    string maxStr = e.Attribute("maxlength") == null ? "" : e.Attribute("maxlength").Value;
                    config.ImportIdentifier = importIdentifier;

                    ImportColumn column = new ImportColumn(e.Attribute("name").Value,
                        e.Attribute("type").Value, maxStr,
                        e.Attribute("required").Value,
                        columnNumber++);

                    config.ImportColumns.Add(column);
                }

            }

            if (configLoaded)
                return config;
            else
                return null;
        }
    }
}
