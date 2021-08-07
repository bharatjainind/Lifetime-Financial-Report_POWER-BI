# Lifetime-Financial-Report_POWER-BI

<img src="https://www.tsg.com/sites/default/files/PowerBI_Banner.png">

It is a Company's Lifetime Financial Analytical Report with **Distorted and Normalised figures** *(for Privacy Reasons and Precautions)* demonstrating a financial performance of a Company engaged in hardware industry specialized in spring manufacturing throughout their Lifetime

<img src="https://i1.wp.com/sqldusty.com/wp-content/uploads/2019/02/Power-BI-architecture-v4.png?ssl=1">

## You can have a look below into the Snapshots of this Report --

<img src="Dark Lifetime Report Snaps/Graphic Dashboard.png">

* As we can observe the overall orders placed within a particular period of time that can be evaluated in terms of Order Counts as well as the Monetary exploitation terms where we can see a Dip in the month of January followed by a Stable rise.
* We can also compare the revenues over time in comparison with the count of orders placed within a month.
* An Order takes time to be executed and goes through multiple phases which we can also observe in the Order Status Funnel tab where 5.63% of orders have been Canceled so far
* France seems to be the Targeted Country for maximum expectation in terms of Sales

<img src="Dark Lifetime Report Snaps/Order Status.png">

We can observe the Status of the Orders demonstrating the health of multiple operations within the Organisation and they can be used to filter the Order Counts as per the Time

<img src="Dark Lifetime Report Snaps/Order Categorized.png">

* France and Poland aeem to be the Leading Countries to make Orders of Springs capturing approximately 75% and 17% Shares of the Targeted Market respectively.
* 90% of the Orders have been Completed So Far meanwhile 3% of the Orders are still Pending to reach their Final Destination.

<img src="Dark Lifetime Report Snaps/Revenue Categorization.png">

* We can observe the Downstream Trend of Revenue generation in the last few months probably due to the Strict Rules over the import of Goods in multiple sections of France as per the improvised provisions for tackling COVID 19 Pendamic Conditions resulting in the revenue to barely cross our Lifetime Average Revenue Mark of 75.25K EURO
* We can also observe the Bifurcation of revenues based on Web Stores where FRANCE constitutes almost 50% of the revenue as a source followed by POLAND Web Store.
* We can Notice the Customer Group Orientation of revenues generated where "Domestic Companies of France" dominating the Customer Group Section

<img src="Dark Lifetime Report Snaps/Tabulation Dashboard.png">

* Here, we can notice the detailed and accurate data organized based on Order Status for Each Country proving the stats of overall collection of revenues throughout its lifetime

<img src="Dark Lifetime Report Snaps/MAP.png">

* Demonstrating the Country performances regarding Financials from an overall perspective

<img src="Dark Lifetime Report Snaps/Branch Analysis.png">

* The Nominal details, as well as Revenue bifercation, can be processed easily by applying relevant FILTERS giving Specific pieces of information about the Customer for future references

## Data Structure & Relationships

<img src="Dark Lifetime Report Snaps/Data Modelling.png">

I also had to Model the Data to build consistent relationships amoung multiple Datasets through M language while Transforming the Data, You can find the necessory scripts below to demostrate How I have ultimately Transformed the data

# For Orders Table

    let
        Source = Csv.Document(File.Contents("C:\Users\asus\Desktop\POWER BI - Project\Raw Vanel Data\Updated Raw Data\1. Orders +.csv"),[Delimiter=",", Columns=48, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"""ID""", type text}, {"Purchase Point", type text}, {"Purchase Date", type datetime}, {"Bill-to Name", type text}, {"Ship-to Name", type text}, {"Grand Total (Base)", type number}, {"Grand Total (Purchased)", type number}, {"Status", type text}, {"Billing Address", type text}, {"Shipping Address", type text}, {"Shipping Information", type text}, {"Customer Email", type text}, {"Customer Group", Int64.Type}, {"Subtotal", type number}, {"Shipping and Handling", type number}, {"Customer Name", type text}, {"Payment Method", type text}, {"Total Refunded", type number}, {"Signifyd Guarantee Decision", type text}, {"Coupon Code", type text}, {"Order Tax", type number}, {"Order Comments", type text}, {"Shipping Method", type text}, {"Subtotal (Base)", type number}, {"Total Due", type number}, {"Total Paid", type number}, {"Subtotal (Purchased)", type number}, {"Weight", type number}, {"Tracking Number", type text}, {"Fax", type text}, {"Region", type text}, {"Postcode", type text}, {"City", type text}, {"Company", type text}, {"Telephone", type text}, {"Country", type text}, {"Fax_1", type text}, {"Region_2", type text}, {"Postcode_3", type text}, {"City_4", type text}, {"Company_5", type text}, {"Telephone_6", type text}, {"Country_7", type text}, {"Product Details", type text}, {"Items Sku", type text}, {"Customer's order number", type text}, {"Country_8", type text}, {"Bar Code", type text}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"""ID""", "order_id"}}),
        #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each [order_id] <> null and [order_id] <> ""),
        #"Removed Duplicates" = Table.Distinct(#"Filtered Rows", {"order_id"}),
        #"Replaced Errors" = Table.ReplaceErrorValues(#"Removed Duplicates", {{"order_id", "N/A"}}),
        #"Inserted Text After Delimiter" = Table.AddColumn(#"Replaced Errors", "main_website_stores", each Text.AfterDelimiter([Purchase Point], " ", 14), type text),
        #"Reordered Columns" = Table.ReorderColumns(#"Inserted Text After Delimiter",{"order_id", "Purchase Point", "main_website_stores", "Purchase Date", "Bill-to Name", "Ship-to Name", "Grand Total (Base)", "Grand Total (Purchased)", "Status", "Billing Address", "Shipping Address", "Shipping Information", "Customer Email", "Customer Group", "Subtotal", "Shipping and Handling", "Customer Name", "Payment Method", "Total Refunded", "Signifyd Guarantee Decision", "Coupon Code", "Order Tax", "Order Comments", "Shipping Method", "Subtotal (Base)", "Total Due", "Total Paid", "Subtotal (Purchased)", "Weight", "Tracking Number", "Fax", "Region", "Postcode", "City", "Company", "Telephone", "Country", "Fax_1", "Region_2", "Postcode_3", "City_4", "Company_5", "Telephone_6", "Country_7", "Product Details", "Items Sku", "Customer's order number", "Country_8", "Bar Code"}),
        #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Purchase Point"}),
        #"Split Column by Delimiter" = Table.SplitColumn(Table.TransformColumnTypes(#"Removed Columns", {{"Purchase Date", type text}}, "en-IN"), "Purchase Date", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"Purchase Date.1", "Purchase Date.2"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Purchase Date.1", type date}, {"Purchase Date.2", type time}}),
        #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"Purchase Date.1", "purchase _date"}, {"Purchase Date.2", "purchase_time"}, {"Bill-to Name", "bill_to_name"}, {"Ship-to Name", "ship_to_name"}}),
        #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns1","","N/A",Replacer.ReplaceValue,{"bill_to_name"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","","N/A",Replacer.ReplaceValue,{"ship_to_name"}),
        #"Renamed Columns2" = Table.RenameColumns(#"Replaced Value1",{{"Grand Total (Base)", "grand_total_(base)"}, {"Grand Total (Purchased)", "grand_total_(purchased)"}, {"Status", "status"}}),
        #"Capitalized Each Word" = Table.TransformColumns(#"Renamed Columns2",{{"status", Text.Proper, type text}}),
        #"Renamed Columns3" = Table.RenameColumns(#"Capitalized Each Word",{{"Billing Address", "billing_address"}}),
        #"Cleaned Text" = Table.TransformColumns(#"Renamed Columns3",{{"billing_address", Text.Clean, type text}}),
        #"Renamed Columns4" = Table.RenameColumns(#"Cleaned Text",{{"Shipping Address", "shipping_address"}}),
        #"Cleaned Text1" = Table.TransformColumns(#"Renamed Columns4",{{"shipping_address", Text.Clean, type text}}),
        #"Replaced Value2" = Table.ReplaceValue(#"Cleaned Text1","Completed","Complete",Replacer.ReplaceValue,{"status"}),
        #"Renamed Columns5" = Table.RenameColumns(#"Replaced Value2",{{"Shipping Information", "shipping_information"}}),
        #"Uppercased Text" = Table.TransformColumns(#"Renamed Columns5",{{"shipping_information", Text.Upper, type text}}),
        #"Trimmed Text" = Table.TransformColumns(#"Uppercased Text",{{"shipping_information", Text.Trim, type text}}),
        #"Replaced Value3" = Table.ReplaceValue(#"Trimmed Text","STANNDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","STRANDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","STANDRAD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value6" = Table.ReplaceValue(#"Replaced Value5","STARDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value7" = Table.ReplaceValue(#"Replaced Value6","STABDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value8" = Table.ReplaceValue(#"Replaced Value7","STANDA5RD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value9" = Table.ReplaceValue(#"Replaced Value8","STANDAQRD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value10" = Table.ReplaceValue(#"Replaced Value9","STANTARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value11" = Table.ReplaceValue(#"Replaced Value10","STANDCARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value12" = Table.ReplaceValue(#"Replaced Value11","STANDAR","STANDARD",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value13" = Table.ReplaceValue(#"Replaced Value12","STANDADR","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value14" = Table.ReplaceValue(#"Replaced Value13","ZPOR GLS","ZPORT GLS",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value15" = Table.ReplaceValue(#"Replaced Value14","ZPORT GLLS","ZPORT GLS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value16" = Table.ReplaceValue(#"Replaced Value15","DHL  - EXPRESS DELIVERY","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value17" = Table.ReplaceValue(#"Replaced Value16","DHL EXPRESSE","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value18" = Table.ReplaceValue(#"Replaced Value17","ZPSECIAL","SPECIAL",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value19" = Table.ReplaceValue(#"Replaced Value18","ZPECIAL","SPECIAL",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value20" = Table.ReplaceValue(#"Replaced Value19","ZPORTGLS","ZPORT GLS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value21" = Table.ReplaceValue(#"Replaced Value20","ZPORRT GLS","ZPORT GLS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value22" = Table.ReplaceValue(#"Replaced Value21","DHL  EXPRESS","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value23" = Table.ReplaceValue(#"Replaced Value22","SPECIAL","ZSPECIAL",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value24" = Table.ReplaceValue(#"Replaced Value23","DHL","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value25" = Table.ReplaceValue(#"Replaced Value24","GLS","GLS EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value26" = Table.ReplaceValue(#"Replaced Value25","CUSTOM NAME","CUSTOM SHIPPING - CUSTOM NAME",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value27" = Table.ReplaceValue(#"Replaced Value26","ZSPECIAL","ZSPECIAL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value28" = Table.ReplaceValue(#"Replaced Value27","ZPORT DHL","ZPORT DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value29" = Table.ReplaceValue(#"Replaced Value28","TRANSPORT DPD","CUSTOM SHIPPING - CUSTOM NAME",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value30" = Table.ReplaceValue(#"Replaced Value29",each [bill_to_name],each if (Text.StartsWith([bill_to_name], " --") or Text.StartsWith([bill_to_name], " - -") or Text.StartsWith([bill_to_name], "--")) then "Not Mentioned" else [bill_to_name],Replacer.ReplaceValue,{"bill_to_name"}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value30",{{"bill_to_name", type text}}),
        #"Replaced Value31" = Table.ReplaceValue(#"Changed Type2",each [ship_to_name],each if (Text.StartsWith([ship_to_name], "-") or Text.StartsWith([ship_to_name], " --") or Text.StartsWith([ship_to_name], "--") or Text.StartsWith([ship_to_name], "  -") or Text.StartsWith([ship_to_name], " - -")) and not Text.Contains([ship_to_name], "a") then "Not Mentioned" else [ship_to_name],Replacer.ReplaceValue,{"ship_to_name"}),
        #"Changed Type3" = Table.TransformColumnTypes(#"Replaced Value31",{{"ship_to_name", type text}}),
        #"Renamed Columns6" = Table.RenameColumns(#"Changed Type3",{{"Customer Email", "customer_email"}, {"Customer Group", "customer_group"}, {"Subtotal", "subtotal"}, {"Shipping and Handling", "shipping_and_handling"}}),
        #"Changed Type4" = Table.TransformColumnTypes(#"Renamed Columns6",{{"customer_group", type text}}),
        #"Replaced Value32" = Table.ReplaceValue(#"Changed Type4","4","INDIVIDUAL",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value33" = Table.ReplaceValue(#"Replaced Value32","1","GENERAL",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value34" = Table.ReplaceValue(#"Replaced Value33","5","DOMESTIC COMPANY",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value35" = Table.ReplaceValue(#"Replaced Value34","6","EU COMPANY",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value36" = Table.ReplaceValue(#"Replaced Value35","7","FOREIGN COMPANY",Replacer.ReplaceText,{"customer_group"}),
        #"Renamed Columns7" = Table.RenameColumns(#"Replaced Value36",{{"Customer Name", "customer_name"}}),
        #"Replaced Value37" = Table.ReplaceValue(#"Renamed Columns7",each [customer_name],each if (Text.StartsWith([customer_name], "-") or Text.StartsWith([customer_name], " --") or Text.StartsWith([customer_name], "--") or Text.StartsWith([customer_name], "  -") or Text.StartsWith([customer_name], " -  -") or Text.StartsWith([customer_name], " - -")) and (not Text.Contains([customer_name], "a") or not Text.Contains([customer_name], "e") or not Text.Contains([customer_name], "i") or not Text.Contains([customer_name], "o") or not Text.Contains([customer_name], "u")) then "Not Mentioned" else [customer_name],Replacer.ReplaceValue,{"customer_name"}),
        #"Changed Type5" = Table.TransformColumnTypes(#"Replaced Value37",{{"customer_name", type text}}),
        #"Renamed Columns8" = Table.RenameColumns(#"Changed Type5",{{"Payment Method", "payment_method"}, {"Total Refunded", "total_refunded"}}),
        #"Replaced Value38" = Table.ReplaceValue(#"Renamed Columns8",null,0,Replacer.ReplaceValue,{"total_refunded"}),
        #"Removed Columns1" = Table.RemoveColumns(#"Replaced Value38",{"Signifyd Guarantee Decision", "Coupon Code", "Order Comments"}),
        #"Renamed Columns9" = Table.RenameColumns(#"Removed Columns1",{{"Order Tax", "order_tax"}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns9",{"Shipping Method"}),
        #"Renamed Columns10" = Table.RenameColumns(#"Removed Columns2",{{"Subtotal (Base)", "subtotal_(base)"}, {"Total Due", "total_due"}, {"Total Paid", "total_paid"}, {"Subtotal (Purchased)", "subtotal_(purchased)"}, {"Weight", "weight"}}),
        #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns10",{"Tracking Number"}),
        #"Renamed Columns11" = Table.RenameColumns(#"Removed Columns3",{{"Fax", "fax"}}),
        #"Removed Columns4" = Table.RemoveColumns(#"Renamed Columns11",{"fax"}),
        #"Replaced Value39" = Table.ReplaceValue(#"Removed Columns4","","NaN",Replacer.ReplaceValue,{"Region"}),
        #"Replaced Value40" = Table.ReplaceValue(#"Replaced Value39","NaN","Not Mentioned",Replacer.ReplaceText,{"Region"}),
        #"Renamed Columns12" = Table.RenameColumns(#"Replaced Value40",{{"Region", "region"}, {"Postcode", "postcode"}, {"City", "city"}}),
        #"Trimmed Text1" = Table.TransformColumns(#"Renamed Columns12",{{"city", Text.Trim, type text}}),
        #"Renamed Columns13" = Table.RenameColumns(#"Trimmed Text1",{{"Company", "company"}}),
        #"Trimmed Text2" = Table.TransformColumns(#"Renamed Columns13",{{"company", Text.Trim, type text}}),
        #"Cleaned Text2" = Table.TransformColumns(#"Trimmed Text2",{{"company", Text.Clean, type text}}),
        #"Replaced Value41" = Table.ReplaceValue(#"Cleaned Text2","","Not Mentioned",Replacer.ReplaceValue,{"company"}),
        #"Renamed Columns14" = Table.RenameColumns(#"Replaced Value41",{{"Telephone", "telephone"}, {"Country", "country"}}),
        #"Removed Columns5" = Table.RemoveColumns(#"Renamed Columns14",{"Fax_1", "Region_2", "Postcode_3", "City_4", "Company_5", "Telephone_6", "Country_7", "Bar Code", "Country_8"}),
        #"Added Index" = Table.AddIndexColumn(#"Removed Columns5", "Index", 0, 1, Int64.Type),
        #"Removed Columns6" = Table.RemoveColumns(#"Added Index",{"Index"}),
        #"Added Index1" = Table.AddIndexColumn(#"Removed Columns6", "Index", 1, 1, Int64.Type),
        #"Inserted Merged Column" = Table.AddColumn(#"Added Index1", "Merged", each Text.Combine({"p_", Text.From([Index], "en-IN")}), type text),
        #"Renamed Columns15" = Table.RenameColumns(#"Inserted Merged Column",{{"Merged", "product_id"}}),
        #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns15",{"order_id", "main_website_stores", "purchase _date", "purchase_time", "bill_to_name", "ship_to_name", "status", "billing_address", "shipping_address", "shipping_information", "customer_email", "customer_group", "shipping_and_handling", "customer_name", "payment_method", "grand_total_(base)", "grand_total_(purchased)", "subtotal", "subtotal_(base)", "subtotal_(purchased)", "total_refunded", "order_tax", "total_due", "total_paid", "Product Details", "weight", "region", "postcode", "city", "company", "telephone", "country", "Items Sku", "Customer's order number", "Index", "product_id"}),
        #"Removed Columns7" = Table.RemoveColumns(#"Reordered Columns1",{"customer_email", "customer_group", "customer_name", "region", "postcode", "city", "company", "telephone", "country", "Product Details", "Items Sku", "Index"}),
        #"Renamed Columns16" = Table.RenameColumns(#"Removed Columns7",{{"order_id", "o.Order ID"}, {"main_website_stores", "o_Main Website Stores"}, {"purchase _date", "o_Purchase Date"}, {"purchase_time", "o_Purchase Time"}, {"bill_to_name", "o_Bill to Name"}, {"ship_to_name", "o_Ship to Name"}, {"status", "o_Status"}, {"billing_address", "o_Billing Address"}, {"shipping_address", "o_Shipping Address"}, {"shipping_information", "o_Shipping Information"}, {"shipping_and_handling", "o_Shipping and Handling"}, {"payment_method", "o_Payment Method"}, {"grand_total_(base)", "o_Grand Total (Base)"}, {"grand_total_(purchased)", "o_Grand Total (Purchased)"}, {"subtotal", "o_Subtotal"}, {"subtotal_(base)", "o_Subtotal (Base)"}, {"subtotal_(purchased)", "o_Subtotal (Purchased)"}, {"total_refunded", "o_Total Refunded"}, {"order_tax", "o_Order Tax"}, {"total_due", "o_Total Due"}, {"total_paid", "o_Total Paid"}, {"weight", "o_Weight"}, {"Customer's order number", "o_Customer Order Number"}, {"product_id", "o_Product ID"}}),
        #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns16",{"o.Order ID", "o_Product ID", "o_Customer Order Number", "o_Purchase Date", "o_Purchase Time", "o_Status", "o_Payment Method", "o_Weight", "o_Order Tax", "o_Shipping and Handling", "o_Subtotal (Base)", "o_Subtotal (Purchased)", "o_Subtotal", "o_Grand Total (Base)", "o_Grand Total (Purchased)", "o_Total Paid", "o_Total Due", "o_Total Refunded", "o_Main Website Stores", "o_Shipping Information", "o_Bill to Name", "o_Ship to Name", "o_Billing Address", "o_Shipping Address"}),
        #"Trimmed Text3" = Table.TransformColumns(#"Reordered Columns2",{{"o_Bill to Name", Text.Trim, type text}, {"o_Ship to Name", Text.Trim, type text}, {"o_Billing Address", Text.Trim, type text}, {"o_Shipping Address", Text.Trim, type text}}),
        #"Cleaned Text3" = Table.TransformColumns(#"Trimmed Text3",{{"o_Bill to Name", Text.Clean, type text}, {"o_Ship to Name", Text.Clean, type text}, {"o_Billing Address",  Text.Clean, type text}, {"o_Shipping Address", Text.Clean, type text}}),
        #"Renamed Columns17" = Table.RenameColumns(#"Cleaned Text3",{{"o_Purchase Date", "Purchase Date"}, {"o_Subtotal (Base)", "o_Subtotal (Base)"}, {"o.Order ID", "o_Order ID"}})
    in
        #"Renamed Columns17"

# For Product Items Ordered Table

    let
        Source = Csv.Document(File.Contents("C:\Users\asus\Desktop\POWER BI - Project\Raw Vanel Data\Updated Raw Data\1. Orders +.csv"),[Delimiter=",", Columns=48, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"""ID""", type text}, {"Purchase Point", type text}, {"Purchase Date", type datetime}, {"Bill-to Name", type text}, {"Ship-to Name", type text}, {"Grand Total (Base)", type number}, {"Grand Total (Purchased)", type number}, {"Status", type text}, {"Billing Address", type text}, {"Shipping Address", type text}, {"Shipping Information", type text}, {"Customer Email", type text}, {"Customer Group", Int64.Type}, {"Subtotal", type number}, {"Shipping and Handling", type number}, {"Customer Name", type text}, {"Payment Method", type text}, {"Total Refunded", type number}, {"Signifyd Guarantee Decision", type text}, {"Coupon Code", type text}, {"Order Tax", type number}, {"Order Comments", type text}, {"Shipping Method", type text}, {"Subtotal (Base)", type number}, {"Total Due", type number}, {"Total Paid", type number}, {"Subtotal (Purchased)", type number}, {"Weight", type number}, {"Tracking Number", type text}, {"Fax", type text}, {"Region", type text}, {"Postcode", type text}, {"City", type text}, {"Company", type text}, {"Telephone", type text}, {"Country", type text}, {"Fax_1", type text}, {"Region_2", type text}, {"Postcode_3", type text}, {"City_4", type text}, {"Company_5", type text}, {"Telephone_6", type text}, {"Country_7", type text}, {"Product Details", type text}, {"Items Sku", type text}, {"Customer's order number", type text}, {"Country_8", type text}, {"Bar Code", type text}}),
        #"Added an Index Column" = Table.AddIndexColumn(#"Changed Type", "Index", 1, 1, Int64.Type),
        #"Inserted Merged Column" = Table.AddColumn(#"Added an Index Column", "product_id", each Text.Combine({"p_", Text.From([Index], "en-IN")}), type text),
        #"Removed the Index Column" = Table.RemoveColumns(#"Inserted Merged Column",{"Index"}),
        #"Removed Other Columns" = Table.SelectColumns(#"Removed the Index Column",{"Product Details", "Customer's order number", "product_id"}),
        #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"Customer's order number", "product_id", "Product Details"}),
        #"Added Index Column as Row Number" = Table.AddIndexColumn(#"Reordered Columns", "Row Number", 1, 1, Int64.Type),
        #"Reordered Columns1" = Table.ReorderColumns(#"Added Index Column as Row Number",{"Customer's order number", "product_id", "Row Number", "Product Details"}),
        #"Split Column by ""|"" Delimiter" = Table.SplitColumn(#"Reordered Columns1", "Product Details", Splitter.SplitTextByDelimiter("|", QuoteStyle.Csv), {"Product Details.1", "Product Details.2", "Product Details.3", "Product Details.4", "Product Details.5", "Product Details.6", "Product Details.7", "Product Details.8", "Product Details.9", "Product Details.10", "Product Details.11", "Product Details.12", "Product Details.13", "Product Details.14", "Product Details.15", "Product Details.16", "Product Details.17", "Product Details.18", "Product Details.19", "Product Details.20", "Product Details.21", "Product Details.22", "Product Details.23", "Product Details.24", "Product Details.25", "Product Details.26", "Product Details.27", "Product Details.28", "Product Details.29", "Product Details.30", "Product Details.31", "Product Details.32", "Product Details.33", "Product Details.34", "Product Details.35", "Product Details.36", "Product Details.37", "Product Details.38", "Product Details.39", "Product Details.40", "Product Details.41", "Product Details.42", "Product Details.43", "Product Details.44", "Product Details.45", "Product Details.46", "Product Details.47", "Product Details.48", "Product Details.49", "Product Details.50", "Product Details.51", "Product Details.52", "Product Details.53", "Product Details.54", "Product Details.55", "Product Details.56", "Product Details.57", "Product Details.58", "Product Details.59", "Product Details.60", "Product Details.61", "Product Details.62", "Product Details.63", "Product Details.64", "Product Details.65", "Product Details.66", "Product Details.67", "Product Details.68", "Product Details.69", "Product Details.70", "Product Details.71", "Product Details.72", "Product Details.73", "Product Details.74", "Product Details.75", "Product Details.76", "Product Details.77", "Product Details.78", "Product Details.79", "Product Details.80", "Product Details.81", "Product Details.82", "Product Details.83", "Product Details.84", "Product Details.85", "Product Details.86", "Product Details.87", "Product Details.88", "Product Details.89", "Product Details.90", "Product Details.91", "Product Details.92", "Product Details.93", "Product Details.94", "Product Details.95", "Product Details.96", "Product Details.97", "Product Details.98", "Product Details.99", "Product Details.100", "Product Details.101", "Product Details.102", "Product Details.103", "Product Details.104", "Product Details.105", "Product Details.106", "Product Details.107", "Product Details.108", "Product Details.109", "Product Details.110", "Product Details.111", "Product Details.112", "Product Details.113", "Product Details.114", "Product Details.115", "Product Details.116", "Product Details.117", "Product Details.118", "Product Details.119", "Product Details.120", "Product Details.121", "Product Details.122", "Product Details.123", "Product Details.124", "Product Details.125", "Product Details.126", "Product Details.127", "Product Details.128", "Product Details.129", "Product Details.130", "Product Details.131", "Product Details.132", "Product Details.133", "Product Details.134", "Product Details.135", "Product Details.136", "Product Details.137", "Product Details.138", "Product Details.139", "Product Details.140", "Product Details.141", "Product Details.142", "Product Details.143", "Product Details.144", "Product Details.145", "Product Details.146", "Product Details.147", "Product Details.148", "Product Details.149", "Product Details.150", "Product Details.151", "Product Details.152", "Product Details.153", "Product Details.154", "Product Details.155", "Product Details.156", "Product Details.157", "Product Details.158", "Product Details.159", "Product Details.160", "Product Details.161", "Product Details.162", "Product Details.163", "Product Details.164", "Product Details.165", "Product Details.166", "Product Details.167", "Product Details.168", "Product Details.169", "Product Details.170", "Product Details.171", "Product Details.172", "Product Details.173", "Product Details.174", "Product Details.175", "Product Details.176", "Product Details.177", "Product Details.178", "Product Details.179", "Product Details.180", "Product Details.181", "Product Details.182", "Product Details.183", "Product Details.184", "Product Details.185", "Product Details.186", "Product Details.187", "Product Details.188", "Product Details.189", "Product Details.190", "Product Details.191", "Product Details.192", "Product Details.193", "Product Details.194", "Product Details.195", "Product Details.196", "Product Details.197", "Product Details.198", "Product Details.199", "Product Details.200", "Product Details.201", "Product Details.202", "Product Details.203", "Product Details.204", "Product Details.205", "Product Details.206", "Product Details.207", "Product Details.208", "Product Details.209", "Product Details.210", "Product Details.211", "Product Details.212", "Product Details.213", "Product Details.214", "Product Details.215", "Product Details.216", "Product Details.217", "Product Details.218", "Product Details.219", "Product Details.220", "Product Details.221", "Product Details.222", "Product Details.223", "Product Details.224", "Product Details.225", "Product Details.226", "Product Details.227", "Product Details.228", "Product Details.229", "Product Details.230", "Product Details.231", "Product Details.232", "Product Details.233", "Product Details.234", "Product Details.235", "Product Details.236", "Product Details.237", "Product Details.238", "Product Details.239", "Product Details.240", "Product Details.241", "Product Details.242", "Product Details.243", "Product Details.244", "Product Details.245", "Product Details.246", "Product Details.247", "Product Details.248", "Product Details.249", "Product Details.250", "Product Details.251", "Product Details.252", "Product Details.253", "Product Details.254", "Product Details.255", "Product Details.256", "Product Details.257", "Product Details.258", "Product Details.259", "Product Details.260", "Product Details.261", "Product Details.262", "Product Details.263", "Product Details.264", "Product Details.265", "Product Details.266", "Product Details.267", "Product Details.268", "Product Details.269", "Product Details.270", "Product Details.271", "Product Details.272", "Product Details.273", "Product Details.274", "Product Details.275", "Product Details.276", "Product Details.277", "Product Details.278", "Product Details.279", "Product Details.280", "Product Details.281", "Product Details.282", "Product Details.283"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by ""|"" Delimiter",{{"Product Details.1", type text}, {"Product Details.2", type text}, {"Product Details.3", type text}, {"Product Details.4", type text}, {"Product Details.5", type text}, {"Product Details.6", type text}, {"Product Details.7", type text}, {"Product Details.8", type text}, {"Product Details.9", type text}, {"Product Details.10", type text}, {"Product Details.11", type text}, {"Product Details.12", type text}, {"Product Details.13", type text}, {"Product Details.14", type text}, {"Product Details.15", type text}, {"Product Details.16", type text}, {"Product Details.17", type text}, {"Product Details.18", type text}, {"Product Details.19", type text}, {"Product Details.20", type text}, {"Product Details.21", type text}, {"Product Details.22", type text}, {"Product Details.23", type text}, {"Product Details.24", type text}, {"Product Details.25", type text}, {"Product Details.26", type text}, {"Product Details.27", type text}, {"Product Details.28", type text}, {"Product Details.29", type text}, {"Product Details.30", type text}, {"Product Details.31", type text}, {"Product Details.32", type text}, {"Product Details.33", type text}, {"Product Details.34", type text}, {"Product Details.35", type text}, {"Product Details.36", type text}, {"Product Details.37", type text}, {"Product Details.38", type text}, {"Product Details.39", type text}, {"Product Details.40", type text}, {"Product Details.41", type text}, {"Product Details.42", type text}, {"Product Details.43", type text}, {"Product Details.44", type text}, {"Product Details.45", type text}, {"Product Details.46", type text}, {"Product Details.47", type text}, {"Product Details.48", type text}, {"Product Details.49", type text}, {"Product Details.50", type text}, {"Product Details.51", type text}, {"Product Details.52", type text}, {"Product Details.53", type text}, {"Product Details.54", type text}, {"Product Details.55", type text}, {"Product Details.56", type text}, {"Product Details.57", type text}, {"Product Details.58", type text}, {"Product Details.59", type text}, {"Product Details.60", type text}, {"Product Details.61", type text}, {"Product Details.62", type text}, {"Product Details.63", type text}, {"Product Details.64", type text}, {"Product Details.65", type text}, {"Product Details.66", type text}, {"Product Details.67", type text}, {"Product Details.68", type text}, {"Product Details.69", type text}, {"Product Details.70", type text}, {"Product Details.71", type text}, {"Product Details.72", type text}, {"Product Details.73", type text}, {"Product Details.74", type text}, {"Product Details.75", type text}, {"Product Details.76", type text}, {"Product Details.77", type text}, {"Product Details.78", type text}, {"Product Details.79", type text}, {"Product Details.80", type text}, {"Product Details.81", type text}, {"Product Details.82", type text}, {"Product Details.83", type text}, {"Product Details.84", type text}, {"Product Details.85", type text}, {"Product Details.86", type text}, {"Product Details.87", type text}, {"Product Details.88", type text}, {"Product Details.89", type text}, {"Product Details.90", type text}, {"Product Details.91", type text}, {"Product Details.92", type text}, {"Product Details.93", type text}, {"Product Details.94", type text}, {"Product Details.95", type text}, {"Product Details.96", type text}, {"Product Details.97", type text}, {"Product Details.98", type text}, {"Product Details.99", type text}, {"Product Details.100", type text}, {"Product Details.101", type text}, {"Product Details.102", type text}, {"Product Details.103", type text}, {"Product Details.104", type text}, {"Product Details.105", type text}, {"Product Details.106", type text}, {"Product Details.107", type text}, {"Product Details.108", type text}, {"Product Details.109", type text}, {"Product Details.110", type text}, {"Product Details.111", type text}, {"Product Details.112", type text}, {"Product Details.113", type text}, {"Product Details.114", type text}, {"Product Details.115", type text}, {"Product Details.116", type text}, {"Product Details.117", type text}, {"Product Details.118", type text}, {"Product Details.119", type text}, {"Product Details.120", type text}, {"Product Details.121", type text}, {"Product Details.122", type text}, {"Product Details.123", type text}, {"Product Details.124", type text}, {"Product Details.125", type text}, {"Product Details.126", type text}, {"Product Details.127", type text}, {"Product Details.128", type text}, {"Product Details.129", type text}, {"Product Details.130", type text}, {"Product Details.131", type text}, {"Product Details.132", type text}, {"Product Details.133", type text}, {"Product Details.134", type text}, {"Product Details.135", type text}, {"Product Details.136", type text}, {"Product Details.137", type text}, {"Product Details.138", type text}, {"Product Details.139", type text}, {"Product Details.140", type text}, {"Product Details.141", type text}, {"Product Details.142", type text}, {"Product Details.143", type text}, {"Product Details.144", type text}, {"Product Details.145", type text}, {"Product Details.146", type text}, {"Product Details.147", type text}, {"Product Details.148", type text}, {"Product Details.149", type text}, {"Product Details.150", type text}, {"Product Details.151", type text}, {"Product Details.152", type text}, {"Product Details.153", type text}, {"Product Details.154", type text}, {"Product Details.155", type text}, {"Product Details.156", type text}, {"Product Details.157", type text}, {"Product Details.158", type text}, {"Product Details.159", type text}, {"Product Details.160", type text}, {"Product Details.161", type text}, {"Product Details.162", type text}, {"Product Details.163", type text}, {"Product Details.164", type text}, {"Product Details.165", type text}, {"Product Details.166", type text}, {"Product Details.167", type text}, {"Product Details.168", type text}, {"Product Details.169", type text}, {"Product Details.170", type text}, {"Product Details.171", type text}, {"Product Details.172", type text}, {"Product Details.173", type text}, {"Product Details.174", type text}, {"Product Details.175", type text}, {"Product Details.176", type text}, {"Product Details.177", type text}, {"Product Details.178", type text}, {"Product Details.179", type text}, {"Product Details.180", type text}, {"Product Details.181", type text}, {"Product Details.182", type text}, {"Product Details.183", type text}, {"Product Details.184", type text}, {"Product Details.185", type text}, {"Product Details.186", type text}, {"Product Details.187", type text}, {"Product Details.188", type text}, {"Product Details.189", type text}, {"Product Details.190", type text}, {"Product Details.191", type text}, {"Product Details.192", type text}, {"Product Details.193", type text}, {"Product Details.194", type text}, {"Product Details.195", type text}, {"Product Details.196", type text}, {"Product Details.197", type text}, {"Product Details.198", type text}, {"Product Details.199", type text}, {"Product Details.200", type text}, {"Product Details.201", type text}, {"Product Details.202", type text}, {"Product Details.203", type text}, {"Product Details.204", type text}, {"Product Details.205", type text}, {"Product Details.206", type text}, {"Product Details.207", type text}, {"Product Details.208", type text}, {"Product Details.209", type text}, {"Product Details.210", type text}, {"Product Details.211", type text}, {"Product Details.212", type text}, {"Product Details.213", type text}, {"Product Details.214", type text}, {"Product Details.215", type text}, {"Product Details.216", type text}, {"Product Details.217", type text}, {"Product Details.218", type text}, {"Product Details.219", type text}, {"Product Details.220", type text}, {"Product Details.221", type text}, {"Product Details.222", type text}, {"Product Details.223", type text}, {"Product Details.224", type text}, {"Product Details.225", type text}, {"Product Details.226", type text}, {"Product Details.227", type text}, {"Product Details.228", type text}, {"Product Details.229", type text}, {"Product Details.230", type text}, {"Product Details.231", type text}, {"Product Details.232", type text}, {"Product Details.233", type text}, {"Product Details.234", type text}, {"Product Details.235", type text}, {"Product Details.236", type text}, {"Product Details.237", type text}, {"Product Details.238", type text}, {"Product Details.239", type text}, {"Product Details.240", type text}, {"Product Details.241", type text}, {"Product Details.242", type text}, {"Product Details.243", type text}, {"Product Details.244", type text}, {"Product Details.245", type text}, {"Product Details.246", type text}, {"Product Details.247", type text}, {"Product Details.248", type text}, {"Product Details.249", type text}, {"Product Details.250", type text}, {"Product Details.251", type text}, {"Product Details.252", type text}, {"Product Details.253", type text}, {"Product Details.254", type text}, {"Product Details.255", type text}, {"Product Details.256", type text}, {"Product Details.257", type text}, {"Product Details.258", type text}, {"Product Details.259", type text}, {"Product Details.260", type text}, {"Product Details.261", type text}, {"Product Details.262", type text}, {"Product Details.263", type text}, {"Product Details.264", type text}, {"Product Details.265", type text}, {"Product Details.266", type text}, {"Product Details.267", type text}, {"Product Details.268", type text}, {"Product Details.269", type text}, {"Product Details.270", type text}, {"Product Details.271", type text}, {"Product Details.272", type text}, {"Product Details.273", type text}, {"Product Details.274", type text}, {"Product Details.275", type text}, {"Product Details.276", type text}, {"Product Details.277", type text}, {"Product Details.278", type text}, {"Product Details.279", type text}, {"Product Details.280", type text}, {"Product Details.281", type text}, {"Product Details.282", type text}, {"Product Details.283", type text}}),
        #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Changed Type1", {"Customer's order number", "product_id", "Row Number"}, "Attribute", "Value"),
        #"Added another Index Column" = Table.AddIndexColumn(#"Unpivoted Other Columns", "Index", 1, 1, Int64.Type),
        #"Added Conditional Column" = Table.AddColumn(#"Added another Index Column", "uni_item", each if Text.StartsWith([Value], "Name") then [Index] else null),
        #"Filled Down" = Table.FillDown(#"Added Conditional Column",{"uni_item"}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Filled Down",{{"uni_item", Int64.Type}}),
        #"Added Unified" = Table.AddColumn(#"Changed Type2", "Unified_Items_under_Product", each [Index]-[uni_item]+1),
        #"Changed Type3" = Table.TransformColumnTypes(#"Added Unified",{{"uni_item", type text}}),
        #"Added product_item_id" = Table.AddColumn(#"Changed Type3", "product_item_id", each [product_id]&"_"&[uni_item]),
        #"Removed Temp Columns" = Table.RemoveColumns(#"Added product_item_id",{"Index", "uni_item", "Unified_Items_under_Product", "Attribute"}),
        #"Reordered Columns2" = Table.ReorderColumns(#"Removed Temp Columns",{"Row Number", "Customer's order number", "product_id", "product_item_id", "Value"}),
        #"Split Column by Delimiter1" = Table.SplitColumn(#"Reordered Columns2", "Value", Splitter.SplitTextByEachDelimiter({":"}, QuoteStyle.Csv, false), {"Value.1", "Value.2"}),
        #"Changed Type4" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"Value.1", type text}, {"Value.2", type text}}),
        #"Pivoted Column" = Table.Pivot(#"Changed Type4", List.Distinct(#"Changed Type4"[Value.1]), "Value.1", "Value.2"),
        #"Sorted Rows" = Table.Sort(#"Pivoted Column",{{"Row Number", Order.Ascending}}),
        #"Removed Columns" = Table.RemoveColumns(#"Sorted Rows",{"Name", "Type"}),
        #"Replaced Value" = Table.ReplaceValue(#"Removed Columns","0%","0",Replacer.ReplaceText,{"Tax Percent"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","23%","0.23",Replacer.ReplaceText,{"Tax Percent"}),
        #"Changed Type5" = Table.TransformColumnTypes(#"Replaced Value1",{{"Tax Percent", type number}}),
        #"Filtered Rows" = Table.SelectRows(#"Changed Type5", each [Tax Percent] <> null and [Tax Percent] <> ""),
        #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows",{"Discount Percent", "Discount Amount (Purchased)", "Discount Amount (Base)", "Thumbnail"}),
        #"Added Conditional Column1" = Table.AddColumn(#"Removed Columns1", "payment_medium", each if Text.StartsWith([#"Price (Purchased)"], "â") then "â" else if Text.StartsWith([#"Price (Purchased)"], "PLN") then "PLN" else null),
        #"Reordered Columns3" = Table.ReorderColumns(#"Added Conditional Column1",{"Row Number", "Customer's order number", "product_id", "product_item_id", "SKU", "payment_medium", "Price (Purchased)", "Price (Base)", "Tax Percent", "Tax Amount (Purchased)", "Tax Amount (Base)", "Qty Ordered", "Qty Available", "Subtotal (Purchased)", "Subtotal (Base)", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded"}),
        #"Changed Type6" = Table.TransformColumnTypes(#"Reordered Columns3",{{"payment_medium", type text}}),
        #"Replaced Value2" = Table.ReplaceValue(#"Changed Type6","â‚¬","",Replacer.ReplaceText,{"Price (Purchased)", "Price (Base)", "Tax Amount (Purchased)", "Tax Amount (Base)", "Subtotal (Purchased)", "Subtotal (Base)"}),
        #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","PLN","",Replacer.ReplaceText,{"Price (Purchased)", "Price (Base)", "Tax Amount (Purchased)", "Tax Amount (Base)", "Subtotal (Purchased)", "Subtotal (Base)"}),
        #"Reordered Columns4" = Table.ReorderColumns(#"Replaced Value3",{"Row Number", "Customer's order number", "product_id", "product_item_id", "SKU", "payment_medium", "Price (Purchased)", "Price (Base)", "Tax Percent", "Tax Amount (Purchased)", "Tax Amount (Base)", "Subtotal (Purchased)", "Subtotal (Base)", "Qty Ordered", "Qty Available", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded"}),
        #"Replaced Value4" = Table.ReplaceValue(#"Reordered Columns4",null,"0",Replacer.ReplaceValue,{"Qty Ordered", "Qty Available", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded"}),
        #"Grouped Rows" = Table.Group(#"Replaced Value4", {"product_id"}, {{"Count", each _, type table [Row Number=number, #"Customer's order number"=nullable text, product_id=text, product_item_id=text, SKU=nullable text, payment_medium=nullable text, #"Price (Purchased)"=nullable text, #"Price (Base)"=nullable text, Tax Percent=nullable number, #"Tax Amount (Purchased)"=nullable text, #"Tax Amount (Base)"=nullable text, #"Subtotal (Purchased)"=nullable text, #"Subtotal (Base)"=nullable text, Qty Ordered=nullable text, Qty Available=nullable text, Qty Invoiced=nullable text, Qty Shipped=nullable text, Qty Canceled=nullable text, Qty Refunded=nullable text]}}),
        #"Added Custom" = Table.AddColumn(#"Grouped Rows", "Custom", each Table.AddIndexColumn([Count], "Index", 1, 1)),
        #"Removed Columns2" = Table.RemoveColumns(#"Added Custom",{"Count"}),
        #"Expanded Custom" = Table.ExpandTableColumn(#"Removed Columns2", "Custom", {"Row Number", "Customer's order number", "product_id", "product_item_id", "SKU", "payment_medium", "Price (Purchased)", "Price (Base)", "Tax Percent", "Tax Amount (Purchased)", "Tax Amount (Base)", "Subtotal (Purchased)", "Subtotal (Base)", "Qty Ordered", "Qty Available", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded", "Index"}, {"Row Number", "Customer's order number", "product_id.1", "product_item_id", "SKU", "payment_medium", "Price (Purchased)", "Price (Base)", "Tax Percent", "Tax Amount (Purchased)", "Tax Amount (Base)", "Subtotal (Purchased)", "Subtotal (Base)", "Qty Ordered", "Qty Available", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded", "Index"}),
        #"Removed Columns3" = Table.RemoveColumns(#"Expanded Custom",{"product_id"}),
        #"Reordered Columns5" = Table.ReorderColumns(#"Removed Columns3",{"Row Number", "Customer's order number", "product_id.1", "product_item_id", "Index", "SKU", "payment_medium", "Price (Purchased)", "Price (Base)", "Tax Percent", "Tax Amount (Purchased)", "Tax Amount (Base)", "Subtotal (Purchased)", "Subtotal (Base)", "Qty Ordered", "Qty Available", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded"}),
        #"Changed Type7" = Table.TransformColumnTypes(#"Reordered Columns5",{{"Row Number", Int64.Type}, {"Customer's order number", type text}, {"product_id.1", type text}, {"product_item_id", type text}, {"Index", Int64.Type}, {"SKU", type text}, {"payment_medium", type text}, {"Price (Purchased)", type number}, {"Price (Base)", type number}, {"Tax Percent", type number}, {"Tax Amount (Purchased)", type number}, {"Tax Amount (Base)", type number}, {"Subtotal (Purchased)", type number}, {"Subtotal (Base)", type number}, {"Qty Ordered", Int64.Type}, {"Qty Available", Int64.Type}, {"Qty Invoiced", Int64.Type}, {"Qty Shipped", Int64.Type}, {"Qty Canceled", Int64.Type}, {"Qty Refunded", Int64.Type}}),
        #"Inserted Merged Column1" = Table.AddColumn(#"Changed Type7", "Merged", each Text.Combine({[product_id.1], "_", Text.From([Index], "en-IN")}), type text),
        #"Reordered Columns6" = Table.ReorderColumns(#"Inserted Merged Column1",{"Row Number", "Customer's order number", "product_id.1", "Merged", "product_item_id", "Index", "SKU", "payment_medium", "Qty Available", "Qty Ordered", "Tax Percent", "Price (Base)", "Tax Amount (Base)", "Subtotal (Base)", "Price (Purchased)", "Tax Amount (Purchased)", "Subtotal (Purchased)", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded"}),
        #"Removed Columns4" = Table.RemoveColumns(#"Reordered Columns6",{"Row Number", "product_item_id", "Index"}),
        #"Renamed Columns" = Table.RenameColumns(#"Removed Columns4",{{"Customer's order number", "Customer Order Number"}, {"product_id.1", "Product ID"}, {"Merged", "Product Item ID"}, {"payment_medium", "Payment Medium"}}),
        #"Added Conditional Column2" = Table.AddColumn(#"Renamed Columns", "Out of Stock", each if [Qty Available] >= [Qty Ordered] then 0 else if [Qty Available] < [Qty Ordered] then 1 else "Error"),
        #"Reordered Columns7" = Table.ReorderColumns(#"Added Conditional Column2",{"Customer Order Number", "Product ID", "Product Item ID", "SKU", "Payment Medium", "Qty Available", "Qty Ordered", "Out of Stock", "Tax Percent", "Price (Base)", "Tax Amount (Base)", "Subtotal (Base)", "Price (Purchased)", "Tax Amount (Purchased)", "Subtotal (Purchased)", "Qty Invoiced", "Qty Shipped", "Qty Canceled", "Qty Refunded"}),
        #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns7",{{"Customer Order Number", "i_Customer Order Number"}, {"Product ID", "i_Product ID"}, {"Product Item ID", "i_Product Item ID"}, {"SKU", "i_SKU"}, {"Payment Medium", "i_Payment Currency"}, {"Qty Available", "i_Qty Available"}, {"Qty Ordered", "i_Qty Ordered"}, {"Out of Stock", "i_Out of Stock"}, {"Tax Percent", "i_Tax Percent"}, {"Price (Base)", "i_Price (Base)"}, {"Tax Amount (Base)", "i_Tax Amount (Base)"}, {"Subtotal (Base)", "i_Subtotal (Base)"}, {"Price (Purchased)", "i_Price (Purchased)"}, {"Tax Amount (Purchased)", "i_Tax Amount (Purchased)"}, {"Subtotal (Purchased)", "i_Subtotal (Purchased)"}, {"Qty Invoiced", "i_Qty Invoiced"}, {"Qty Shipped", "i_Qty Shipped"}, {"Qty Canceled", "i_Qty Canceled"}, {"Qty Refunded", "i_Qty Refunded"}})
    in
        #"Renamed Columns1"

# For Cusomers Ordered Table

    let
        Source = Csv.Document(File.Contents("C:\Users\asus\Desktop\POWER BI - Project\Raw Vanel Data\Updated Raw Data\1. Orders +.csv"),[Delimiter=",", Columns=48, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"""ID""", type text}, {"Purchase Point", type text}, {"Purchase Date", type datetime}, {"Bill-to Name", type text}, {"Ship-to Name", type text}, {"Grand Total (Base)", type number}, {"Grand Total (Purchased)", type number}, {"Status", type text}, {"Billing Address", type text}, {"Shipping Address", type text}, {"Shipping Information", type text}, {"Customer Email", type text}, {"Customer Group", Int64.Type}, {"Subtotal", type number}, {"Shipping and Handling", type number}, {"Customer Name", type text}, {"Payment Method", type text}, {"Total Refunded", type number}, {"Signifyd Guarantee Decision", type text}, {"Coupon Code", type text}, {"Order Tax", type number}, {"Order Comments", type text}, {"Shipping Method", type text}, {"Subtotal (Base)", type number}, {"Total Due", type number}, {"Total Paid", type number}, {"Subtotal (Purchased)", type number}, {"Weight", type number}, {"Tracking Number", type text}, {"Fax", type text}, {"Region", type text}, {"Postcode", type text}, {"City", type text}, {"Company", type text}, {"Telephone", type text}, {"Country", type text}, {"Fax_1", type text}, {"Region_2", type text}, {"Postcode_3", type text}, {"City_4", type text}, {"Company_5", type text}, {"Telephone_6", type text}, {"Country_7", type text}, {"Product Details", type text}, {"Items Sku", type text}, {"Customer's order number", type text}, {"Country_8", type text}, {"Bar Code", type text}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"""ID""", "order_id"}}),
        #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each [order_id] <> null and [order_id] <> ""),
        #"Removed Duplicates" = Table.Distinct(#"Filtered Rows", {"order_id"}),
        #"Replaced Errors" = Table.ReplaceErrorValues(#"Removed Duplicates", {{"order_id", "N/A"}}),
        #"Inserted Text After Delimiter" = Table.AddColumn(#"Replaced Errors", "main_website_stores", each Text.AfterDelimiter([Purchase Point], " ", 14), type text),
        #"Reordered Columns" = Table.ReorderColumns(#"Inserted Text After Delimiter",{"order_id", "Purchase Point", "main_website_stores", "Purchase Date", "Bill-to Name", "Ship-to Name", "Grand Total (Base)", "Grand Total (Purchased)", "Status", "Billing Address", "Shipping Address", "Shipping Information", "Customer Email", "Customer Group", "Subtotal", "Shipping and Handling", "Customer Name", "Payment Method", "Total Refunded", "Signifyd Guarantee Decision", "Coupon Code", "Order Tax", "Order Comments", "Shipping Method", "Subtotal (Base)", "Total Due", "Total Paid", "Subtotal (Purchased)", "Weight", "Tracking Number", "Fax", "Region", "Postcode", "City", "Company", "Telephone", "Country", "Fax_1", "Region_2", "Postcode_3", "City_4", "Company_5", "Telephone_6", "Country_7", "Product Details", "Items Sku", "Customer's order number", "Country_8", "Bar Code"}),
        #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Purchase Point"}),
        #"Split Column by Delimiter" = Table.SplitColumn(Table.TransformColumnTypes(#"Removed Columns", {{"Purchase Date", type text}}, "en-IN"), "Purchase Date", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"Purchase Date.1", "Purchase Date.2"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Purchase Date.1", type date}, {"Purchase Date.2", type time}}),
        #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"Purchase Date.1", "purchase _date"}, {"Purchase Date.2", "purchase_time"}, {"Bill-to Name", "bill_to_name"}, {"Ship-to Name", "ship_to_name"}}),
        #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns1","","N/A",Replacer.ReplaceValue,{"bill_to_name"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","","N/A",Replacer.ReplaceValue,{"ship_to_name"}),
        #"Renamed Columns2" = Table.RenameColumns(#"Replaced Value1",{{"Grand Total (Base)", "grand_total_(base)"}, {"Grand Total (Purchased)", "grand_total_(purchased)"}, {"Status", "status"}}),
        #"Capitalized Each Word" = Table.TransformColumns(#"Renamed Columns2",{{"status", Text.Proper, type text}}),
        #"Renamed Columns3" = Table.RenameColumns(#"Capitalized Each Word",{{"Billing Address", "billing_address"}}),
        #"Cleaned Text" = Table.TransformColumns(#"Renamed Columns3",{{"billing_address", Text.Clean, type text}}),
        #"Renamed Columns4" = Table.RenameColumns(#"Cleaned Text",{{"Shipping Address", "shipping_address"}}),
        #"Cleaned Text1" = Table.TransformColumns(#"Renamed Columns4",{{"shipping_address", Text.Clean, type text}}),
        #"Replaced Value2" = Table.ReplaceValue(#"Cleaned Text1","Completed","Complete",Replacer.ReplaceValue,{"status"}),
        #"Renamed Columns5" = Table.RenameColumns(#"Replaced Value2",{{"Shipping Information", "shipping_information"}}),
        #"Uppercased Text" = Table.TransformColumns(#"Renamed Columns5",{{"shipping_information", Text.Upper, type text}}),
        #"Trimmed Text" = Table.TransformColumns(#"Uppercased Text",{{"shipping_information", Text.Trim, type text}}),
        #"Replaced Value3" = Table.ReplaceValue(#"Trimmed Text","STANNDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","STRANDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","STANDRAD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value6" = Table.ReplaceValue(#"Replaced Value5","STARDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value7" = Table.ReplaceValue(#"Replaced Value6","STABDARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value8" = Table.ReplaceValue(#"Replaced Value7","STANDA5RD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value9" = Table.ReplaceValue(#"Replaced Value8","STANDAQRD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value10" = Table.ReplaceValue(#"Replaced Value9","STANTARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value11" = Table.ReplaceValue(#"Replaced Value10","STANDCARD","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value12" = Table.ReplaceValue(#"Replaced Value11","STANDAR","STANDARD",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value13" = Table.ReplaceValue(#"Replaced Value12","STANDADR","STANDARD",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value14" = Table.ReplaceValue(#"Replaced Value13","ZPOR GLS","ZPORT GLS",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value15" = Table.ReplaceValue(#"Replaced Value14","ZPORT GLLS","ZPORT GLS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value16" = Table.ReplaceValue(#"Replaced Value15","DHL  - EXPRESS DELIVERY","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value17" = Table.ReplaceValue(#"Replaced Value16","DHL EXPRESSE","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value18" = Table.ReplaceValue(#"Replaced Value17","ZPSECIAL","SPECIAL",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value19" = Table.ReplaceValue(#"Replaced Value18","ZPECIAL","SPECIAL",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value20" = Table.ReplaceValue(#"Replaced Value19","ZPORTGLS","ZPORT GLS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value21" = Table.ReplaceValue(#"Replaced Value20","ZPORRT GLS","ZPORT GLS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value22" = Table.ReplaceValue(#"Replaced Value21","DHL  EXPRESS","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value23" = Table.ReplaceValue(#"Replaced Value22","SPECIAL","ZSPECIAL",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value24" = Table.ReplaceValue(#"Replaced Value23","DHL","DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value25" = Table.ReplaceValue(#"Replaced Value24","GLS","GLS EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value26" = Table.ReplaceValue(#"Replaced Value25","CUSTOM NAME","CUSTOM SHIPPING - CUSTOM NAME",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value27" = Table.ReplaceValue(#"Replaced Value26","ZSPECIAL","ZSPECIAL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value28" = Table.ReplaceValue(#"Replaced Value27","ZPORT DHL","ZPORT DHL EXPRESS",Replacer.ReplaceValue,{"shipping_information"}),
        #"Replaced Value29" = Table.ReplaceValue(#"Replaced Value28","TRANSPORT DPD","CUSTOM SHIPPING - CUSTOM NAME",Replacer.ReplaceText,{"shipping_information"}),
        #"Replaced Value30" = Table.ReplaceValue(#"Replaced Value29",each [bill_to_name],each if (Text.StartsWith([bill_to_name], " --") or Text.StartsWith([bill_to_name], " - -") or Text.StartsWith([bill_to_name], "--")) then "Not Mentioned" else [bill_to_name],Replacer.ReplaceValue,{"bill_to_name"}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value30",{{"bill_to_name", type text}}),
        #"Replaced Value31" = Table.ReplaceValue(#"Changed Type2",each [ship_to_name],each if (Text.StartsWith([ship_to_name], "-") or Text.StartsWith([ship_to_name], " --") or Text.StartsWith([ship_to_name], "--") or Text.StartsWith([ship_to_name], "  -") or Text.StartsWith([ship_to_name], " - -")) and not Text.Contains([ship_to_name], "a") then "Not Mentioned" else [ship_to_name],Replacer.ReplaceValue,{"ship_to_name"}),
        #"Changed Type3" = Table.TransformColumnTypes(#"Replaced Value31",{{"ship_to_name", type text}}),
        #"Renamed Columns6" = Table.RenameColumns(#"Changed Type3",{{"Customer Email", "customer_email"}, {"Customer Group", "customer_group"}, {"Subtotal", "subtotal"}, {"Shipping and Handling", "shipping_and_handling"}}),
        #"Changed Type4" = Table.TransformColumnTypes(#"Renamed Columns6",{{"customer_group", type text}}),
        #"Replaced Value32" = Table.ReplaceValue(#"Changed Type4","4","INDIVIDUAL",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value33" = Table.ReplaceValue(#"Replaced Value32","1","GENERAL",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value34" = Table.ReplaceValue(#"Replaced Value33","5","DOMESTIC COMPANY",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value35" = Table.ReplaceValue(#"Replaced Value34","6","EU COMPANY",Replacer.ReplaceText,{"customer_group"}),
        #"Replaced Value36" = Table.ReplaceValue(#"Replaced Value35","7","FOREIGN COMPANY",Replacer.ReplaceText,{"customer_group"}),
        #"Renamed Columns7" = Table.RenameColumns(#"Replaced Value36",{{"Customer Name", "customer_name"}}),
        #"Replaced Value37" = Table.ReplaceValue(#"Renamed Columns7",each [customer_name],each if (Text.StartsWith([customer_name], "-") or Text.StartsWith([customer_name], " --") or Text.StartsWith([customer_name], "--") or Text.StartsWith([customer_name], "  -") or Text.StartsWith([customer_name], " -  -") or Text.StartsWith([customer_name], " - -")) and (not Text.Contains([customer_name], "a") or not Text.Contains([customer_name], "e") or not Text.Contains([customer_name], "i") or not Text.Contains([customer_name], "o") or not Text.Contains([customer_name], "u")) then "Not Mentioned" else [customer_name],Replacer.ReplaceValue,{"customer_name"}),
        #"Changed Type5" = Table.TransformColumnTypes(#"Replaced Value37",{{"customer_name", type text}}),
        #"Renamed Columns8" = Table.RenameColumns(#"Changed Type5",{{"Payment Method", "payment_method"}, {"Total Refunded", "total_refunded"}}),
        #"Replaced Value38" = Table.ReplaceValue(#"Renamed Columns8",null,0,Replacer.ReplaceValue,{"total_refunded"}),
        #"Removed Columns1" = Table.RemoveColumns(#"Replaced Value38",{"Signifyd Guarantee Decision", "Coupon Code", "Order Comments"}),
        #"Renamed Columns9" = Table.RenameColumns(#"Removed Columns1",{{"Order Tax", "order_tax"}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns9",{"Shipping Method"}),
        #"Renamed Columns10" = Table.RenameColumns(#"Removed Columns2",{{"Subtotal (Base)", "subtotal_(base)"}, {"Total Due", "total_due"}, {"Total Paid", "total_paid"}, {"Subtotal (Purchased)", "subtotal_(purchased)"}, {"Weight", "weight"}}),
        #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns10",{"Tracking Number"}),
        #"Renamed Columns11" = Table.RenameColumns(#"Removed Columns3",{{"Fax", "fax"}}),
        #"Removed Columns4" = Table.RemoveColumns(#"Renamed Columns11",{"fax"}),
        #"Replaced Value39" = Table.ReplaceValue(#"Removed Columns4","","NaN",Replacer.ReplaceValue,{"Region"}),
        #"Replaced Value40" = Table.ReplaceValue(#"Replaced Value39","NaN","Not Mentioned",Replacer.ReplaceText,{"Region"}),
        #"Renamed Columns12" = Table.RenameColumns(#"Replaced Value40",{{"Region", "region"}, {"Postcode", "postcode"}, {"City", "city"}}),
        #"Trimmed Text1" = Table.TransformColumns(#"Renamed Columns12",{{"city", Text.Trim, type text}}),
        #"Renamed Columns13" = Table.RenameColumns(#"Trimmed Text1",{{"Company", "company"}}),
        #"Trimmed Text2" = Table.TransformColumns(#"Renamed Columns13",{{"company", Text.Trim, type text}}),
        #"Cleaned Text2" = Table.TransformColumns(#"Trimmed Text2",{{"company", Text.Clean, type text}}),
        #"Replaced Value41" = Table.ReplaceValue(#"Cleaned Text2","","Not Mentioned",Replacer.ReplaceValue,{"company"}),
        #"Renamed Columns14" = Table.RenameColumns(#"Replaced Value41",{{"Telephone", "telephone"}, {"Country", "country"}}),
        #"Removed Columns5" = Table.RemoveColumns(#"Renamed Columns14",{"Fax_1", "Region_2", "Postcode_3", "City_4", "Company_5", "Telephone_6", "Country_7", "Bar Code", "Country_8"}),
        #"Added Index" = Table.AddIndexColumn(#"Removed Columns5", "Index", 0, 1, Int64.Type),
        #"Removed Columns6" = Table.RemoveColumns(#"Added Index",{"Index"}),
        #"Added Index1" = Table.AddIndexColumn(#"Removed Columns6", "Index", 1, 1, Int64.Type),
        #"Inserted Merged Column" = Table.AddColumn(#"Added Index1", "Merged", each Text.Combine({"p_", Text.From([Index], "en-IN")}), type text),
        #"Renamed Columns15" = Table.RenameColumns(#"Inserted Merged Column",{{"Merged", "product_id"}}),
        #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns15",{"order_id", "main_website_stores", "purchase _date", "purchase_time", "bill_to_name", "ship_to_name", "status", "billing_address", "shipping_address", "shipping_information", "customer_email", "customer_group", "shipping_and_handling", "customer_name", "payment_method", "grand_total_(base)", "grand_total_(purchased)", "subtotal", "subtotal_(base)", "subtotal_(purchased)", "total_refunded", "order_tax", "total_due", "total_paid", "Product Details", "weight", "region", "postcode", "city", "company", "telephone", "country", "Items Sku", "Customer's order number", "Index", "product_id"}),
        #"Removed Other Columns" = Table.SelectColumns(#"Reordered Columns1",{"order_id", "shipping_address", "customer_email", "customer_group", "customer_name", "region", "postcode", "city", "company", "telephone", "country", "Customer's order number", "product_id"}),
        #"Renamed Columns16" = Table.RenameColumns(#"Removed Other Columns",{{"order_id", "c_Order ID"}, {"shipping_address", "c_Address"}, {"customer_email", "c_Customer Email"}, {"customer_group", "c_Customer Group"}, {"customer_name", "c_Customer Name"}, {"region", "c_Region"}, {"postcode", "c_Postcode"}, {"city", "c_City"}, {"company", "c_Company"}, {"telephone", "c_Telephone"}, {"country", "c_Country"}, {"Customer's order number", "c_Customer Order Number"}, {"product_id", "c_Product ID"}}),
        #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns16",{"c_Order ID", "c_Customer Order Number", "c_Product ID", "c_Customer Group", "c_Customer Email", "c_Customer Name", "c_Company", "c_Telephone", "c_Region", "c_Postcode", "c_City", "c_Address", "c_Country"}),
        #"Removed Columns7" = Table.RemoveColumns(#"Reordered Columns2",{"c_Product ID"})
    in
        #"Removed Columns7"

# For Proforma Table

    let
        Source = Csv.Document(File.Contents("C:\Users\asus\Desktop\POWER BI - Project\Raw Vanel Data\Updated Raw Data\2. Proforma Invoices +.csv"),[Delimiter=",", Columns=5, Encoding=1252, QuoteStyle=QuoteStyle.None]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Proforma ID", Int64.Type}, {"Order ID", type text}, {"Created", type datetime}, {"Client e-mail", type text}, {"Is Send Proforma", Int64.Type}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Proforma ID", "p_Proforma ID"}, {"Order ID", "p_Order ID"}, {"Client e-mail", "p_Customer E-Mail"}}),
        #"Inserted Date" = Table.AddColumn(#"Renamed Columns", "Date", each DateTime.Date([Created]), type date),
        #"Inserted Time" = Table.AddColumn(#"Inserted Date", "Time", each DateTime.Time([Created]), type time),
        #"Reordered Columns" = Table.ReorderColumns(#"Inserted Time",{"p_Proforma ID", "p_Order ID", "Created", "Date", "Time", "p_Customer E-Mail", "Is Send Proforma"}),
        #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Created"}),
        #"Trimmed Text" = Table.TransformColumns(#"Removed Columns",{{"p_Customer E-Mail", Text.Trim, type text}}),
        #"Renamed Columns1" = Table.RenameColumns(#"Trimmed Text",{{"Date", "p_Date"}, {"Time", "p_Time"}})
    in
        #"Renamed Columns1"

# For Invoices Table

    let
        Source = Csv.Document(File.Contents("C:\Users\asus\Desktop\POWER BI - Project\Raw Vanel Data\Updated Raw Data\3. Invoices +.csv"),[Delimiter=",", Columns=18, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Invoice", type text}, {"Invoice Date", type datetime}, {"Order #", type text}, {"Order Date", type datetime}, {"Bill-to Name", type text}, {"Status", type text}, {"Grand Total (Base)", type number}, {"Grand Total (Purchased)", type number}, {"Purchased From", type text}, {"Billing Address", type text}, {"Shipping Address", type text}, {"Customer Name", type text}, {"Email", type text}, {"Customer Group", type text}, {"Payment Method", type text}, {"Shipping Information", type text}, {"Subtotal", type number}, {"Shipping and Handling", type number}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Order #", "in_Order ID"}, {"Invoice", "in_Invoice ID"}}),
        #"Inserted Date" = Table.AddColumn(#"Renamed Columns", "Date", each DateTime.Date([Invoice Date]), type date),
        #"Inserted Time" = Table.AddColumn(#"Inserted Date", "Time", each DateTime.Time([Invoice Date]), type time),
        #"Renamed Columns1" = Table.RenameColumns(#"Inserted Time",{{"Date", "in_Invoice Date"}, {"Time", "in_Invoice Time"}}),
        #"Inserted Date1" = Table.AddColumn(#"Renamed Columns1", "Date", each DateTime.Date([Order Date]), type date),
        #"Renamed Columns2" = Table.RenameColumns(#"Inserted Date1",{{"Date", "in_Order Date"}}),
        #"Inserted Time1" = Table.AddColumn(#"Renamed Columns2", "Time", each DateTime.Time([Order Date]), type time),
        #"Renamed Columns3" = Table.RenameColumns(#"Inserted Time1",{{"Time", "in_Order Time"}}),
        #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns3",{"in_Invoice ID", "Invoice Date", "in_Invoice Time", "in_Invoice Date", "in_Order ID", "Order Date", "in_Order Date", "in_Order Time", "Bill-to Name", "Status", "Grand Total (Base)", "Grand Total (Purchased)", "Purchased From", "Billing Address", "Shipping Address", "Customer Name", "Email", "Customer Group", "Payment Method", "Shipping Information", "Subtotal", "Shipping and Handling"}),
        #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Invoice Date", "Order Date", "Purchased From", "Billing Address", "Shipping Address"}),
        #"Renamed Columns4" = Table.RenameColumns(#"Removed Columns",{{"Customer Group", "in_Customer Group"}, {"Email", "in_Email"}}),
        #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns4",{"Payment Method", "Shipping Information"}),
        #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns1",{"in_Invoice ID", "in_Invoice Time", "in_Invoice Date", "in_Order ID", "in_Order Date", "in_Order Time", "Bill-to Name", "Status", "Customer Name", "in_Email", "in_Customer Group", "Subtotal", "Shipping and Handling", "Grand Total (Base)", "Grand Total (Purchased)"}),
        #"Renamed Columns5" = Table.RenameColumns(#"Reordered Columns1",{{"Status", "in_Status"}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns5",{"Bill-to Name", "Customer Name"}),
        #"Reordered Columns2" = Table.ReorderColumns(#"Removed Columns2",{"in_Invoice ID", "in_Invoice Date", "in_Invoice Time", "in_Order ID", "in_Order Date", "in_Order Time", "in_Status", "in_Email", "in_Customer Group", "Shipping and Handling", "Subtotal", "Grand Total (Base)", "Grand Total (Purchased)"})
    in
        #"Reordered Columns2"

# For Shipments Table

    let
        Source = Csv.Document(File.Contents("C:\Users\asus\Desktop\POWER BI - Project\Raw Vanel Data\Updated Raw Data\4. Shipments +.csv"),[Delimiter=",", Columns=16, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Shipment", type text}, {"Ship Date", type datetime}, {"Order #", type text}, {"Order Date", type datetime}, {"Ship-to Name", type text}, {"Total Quantity", Int64.Type}, {"Order Status", type text}, {"Purchased From", type text}, {"Customer Name", type text}, {"Email", type text}, {"Customer Group", Int64.Type}, {"Billing Address", type text}, {"Shipping Address", type text}, {"Payment Method", type text}, {"Shipping Information", type text}, {"Shipment Status", Int64.Type}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Shipment", "s_Shipment ID"}, {"Order #", "s_Order ID"}}),
        #"Inserted Date" = Table.AddColumn(#"Renamed Columns", "Date", each DateTime.Date([Ship Date]), type date),
        #"Renamed Columns1" = Table.RenameColumns(#"Inserted Date",{{"Date", "s_Shipment Date"}}),
        #"Inserted Time" = Table.AddColumn(#"Renamed Columns1", "Time", each DateTime.Time([Ship Date]), type time),
        #"Renamed Columns2" = Table.RenameColumns(#"Inserted Time",{{"Time", "s_Shipment Time"}}),
        #"Inserted Date1" = Table.AddColumn(#"Renamed Columns2", "Date", each DateTime.Date([Order Date]), type date),
        #"Inserted Time1" = Table.AddColumn(#"Inserted Date1", "Time", each DateTime.Time([Order Date]), type time),
        #"Renamed Columns3" = Table.RenameColumns(#"Inserted Time1",{{"Date", "s_Order Date"}, {"Time", "s_Order Time"}}),
        #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns3",{"s_Shipment ID", "Ship Date", "s_Order ID", "Order Date", "s_Shipment Date", "s_Shipment Time", "s_Order Date", "s_Order Time", "Ship-to Name", "Total Quantity", "Order Status", "Purchased From", "Customer Name", "Email", "Customer Group", "Billing Address", "Shipping Address", "Payment Method", "Shipping Information", "Shipment Status"}),
        #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Ship Date", "Order Date"}),
        #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns",{"s_Shipment ID", "s_Shipment Date", "s_Shipment Time", "s_Order ID", "s_Order Date", "s_Order Time", "Ship-to Name", "Total Quantity", "Order Status", "Purchased From", "Customer Name", "Email", "Customer Group", "Billing Address", "Shipping Address", "Payment Method", "Shipping Information", "Shipment Status"}),
        #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns1",{"Ship-to Name"}),
        #"Renamed Columns4" = Table.RenameColumns(#"Removed Columns1",{{"Total Quantity", "s_Total Quantity"}, {"Order Status", "s_Order Status"}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns4",{"Purchased From", "Customer Name"}),
        #"Renamed Columns5" = Table.RenameColumns(#"Removed Columns2",{{"Email", "s_Email"}}),
        #"Trimmed Text" = Table.TransformColumns(#"Renamed Columns5",{{"s_Email", Text.Trim, type text}}),
        #"Removed Columns3" = Table.RemoveColumns(#"Trimmed Text",{"Billing Address", "Shipping Address", "Payment Method"}),
        #"Capitalized Each Word" = Table.TransformColumns(#"Removed Columns3",{{"s_Order Status", Text.Proper, type text}}),
        #"Removed Columns4" = Table.RemoveColumns(#"Capitalized Each Word",{"Shipping Information"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Removed Columns4",{{"Customer Group", type text}}),
        #"Replaced Value32" = Table.ReplaceValue(#"Changed Type1","4","INDIVIDUAL",Replacer.ReplaceText,{"Customer Group"}),
        #"Replaced Value33" = Table.ReplaceValue(#"Replaced Value32","1","GENERAL",Replacer.ReplaceText,{"Customer Group"}),
        #"Replaced Value34" = Table.ReplaceValue(#"Replaced Value33","5","DOMESTIC COMPANY",Replacer.ReplaceText,{"Customer Group"}),
        #"Replaced Value35" = Table.ReplaceValue(#"Replaced Value34","6","EU COMPANY",Replacer.ReplaceText,{"Customer Group"}),
        #"Replaced Value36" = Table.ReplaceValue(#"Replaced Value35","7","FOREIGN COMPANY",Replacer.ReplaceText,{"Customer Group"}),
        #"Renamed Columns6" = Table.RenameColumns(#"Replaced Value36",{{"Customer Group", "s_Customer Group"}}),
        #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns6",{"s_Shipment ID", "s_Shipment Date", "s_Shipment Time", "s_Order ID", "s_Order Date", "s_Order Time", "s_Email", "s_Customer Group", "s_Order Status", "s_Total Quantity", "Shipment Status"}),
        #"Renamed Columns7" = Table.RenameColumns(#"Reordered Columns2",{{"Shipment Status", "s_Shipment Status"}}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Columns7",{{"s_Shipment Status", type text}}),
        #"Replaced Value" = Table.ReplaceValue(#"Changed Type2",null,"Pending",Replacer.ReplaceValue,{"s_Shipment Status"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","0","Completed",Replacer.ReplaceText,{"s_Shipment Status"}),
        #"Reordered Columns3" = Table.ReorderColumns(#"Replaced Value1",{"s_Shipment ID", "s_Shipment Date", "s_Shipment Time", "s_Order ID", "s_Order Date", "s_Order Time", "s_Email", "s_Customer Group", "s_Order Status", "s_Shipment Status", "s_Total Quantity"})
    in
        #"Reordered Columns3"
        
* *This Project Report is only meant for Analytical Purpose and nothing more than that.
* *Please Visit the real Dark Lifetime Report.pbix in this repository for full dasboard experience.
