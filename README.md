# Various tool for finance

### Degressivity Solver with Python
This program has a specific yet valuable use on a Structured Products desk. It allows for quick calculation of the period-by-period degressivity of the autocall barrier based on the product's parameters.

### Download Historical Data with VBA Using Bloomberg API
A connection to a Bloomberg terminal is required to run the VBA code embedded in the Excel file. This code enables the easy and much faster download of historical data for multiple assets (initially prices, but also volumes or any other historical data) compared to using the '=BDH' formulas, which cannot be automated with VBA.
You simply need to enter the tickers in column A and run the script using the button available on the page. Additional buttons are provided to download historical data for the components of various indices. Please note that the tickers must follow the Bloomberg format, such as 'AAPL UN Equity', 'CAC Index', or 'XAU Curncy'.
