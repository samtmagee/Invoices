$body = "<!DOCTYPE html>
<html>

<head>
    <style>
        body {
            background-color: white;
        }

        h1.Title {
            color: black;
            text-align: right;
        }

        p {
            color: black;
        }

        div.Canvas {
            background-color: aqua;
            width: 100%;
            float: left;

        }

        div.Logo-Title {
            background-color: green;
            width: 100%;
            float: left;
        }

        div.Logo {
            background-color: rgb(98, 95, 160);
            width: 50%;
            float: left;
        }

        div.Title {
            background-color: blue;
            width: 50%;
            text-align: right;
            float: left;
        }

        div.Addresses {
            background-color: grey;
            width: 100%;
            float: left;
        }

        div.customerAddress {
            background-color: cadetblue;
            width: 50%;
            float: left;
        }

        div.ourAddress {
            background-color: rgb(255, 0, 0);
            width: 50%;
            float: left;
        }

        div.Date-Statement-Customer {
            background-color: rgb(144, 173, 141);
            width: 100%;
            float: left;
            text-align: right;
        }

        div.textDeclaraction {
            background-color: rgb(173, 156, 141);
            width: 100%;
            float: left;
            text-align: center;
            font-weight: bold;

        }

        div.Page {
            background-color: rgb(153, 141, 173);
            width: 100%;
            float: left;
            text-align: right;
        }
        @media print {
            .pagebreak { page-break-before: always; } /* page-break-after works, as well */
        }
    </style>
</head>

<body>
    <div class='Canvas'>
        <div class='Logo-Title'>
            <div class='Logo'><img src='C:\GITHub\Invoices\Images\Logo.jpg' width='100px'></div>
            <div class='Title'>
                <h1>Tax Invoice / Statement</h1>
            </div>
        </div>
        <div class='Addresses'>
            <div class='customerAddress'>
                <p>Customer Line 1</p>
                <p>Customer Line 2</p>
                <p>Customer Line 3</p>
            </div>
            <div class='ourAddress'>
                <p>Our Address Line 1</p>
                <p>Our Address Line 2</p>
                <p>Our Address Line 3</p>
            </div>
        </div>
        <div class='Date-Statement-Customer'>
            <p>Date: </p>
            <p>Statement #: </p>
            <p>Customer #: </p>
        </div>
        <div class='Table'>"
        
        
        foreach ($item in $export) {
            $body += "<p>Name: $($item.'Name')</p>
                        <p>Description: $($item.'Description')</p>
                        <p>Line Cost: $($item.'Line Cost')</p>
                        
                        <div class='pagebreak'> </div>"
        }        
        
        $body+= "</div>
        <div class='textDeclaraction'>
            <p>This amount will be paid via direct debit on or around the 15th of MMMM YY</p>
            <p>Detailed transaction listing can be found in the accompanying document BP Transaction Listing.</p>
            <p>Invoice queries should be direct to XXXX XXX XXXX</p>
        </div>
        <div class='Page'>
            <p>Page X of Y</p>
        </div>
    </div>
</body>

</html>"