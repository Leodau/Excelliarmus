// Excelliarmus - Your tiny excel generator.
// Michele LEO 2018
// github.com/leodau

let buttonD = document.getElementsByTagName("button")[3];

buttonD.onclick = (event) => {
	let excel = new Excelliarmus();
	let sheet = excel.addSheet("Styles and Formats");

	// Create simple styling rules, for the background fill.
	let bgRed = excel.addStyle({backgroundColor: "#ff3300"});
	let bgBlue = excel.addStyle({backgroundColor: "#0099ff"});
	let bgGreen = excel.addStyle({backgroundColor: "#33cc33"});

	// Apply styles for each insert.
	excel.insert({column: 1, row: 2, value: "#ff3300 Background", style: bgRed});
	excel.insert({column: 1, row: 3, value: "#0099ff Background", style: bgBlue});
	excel.insert({column: 1, row: 4, value: "#33cc33 Background", style: bgGreen});

	// Font styling.
	let fontBig = excel.addStyle({fontSize: "20"});
	let fontSmall = excel.addStyle({fontSize: "8"});
	let fontBold = excel.addStyle({fontStyle: "bold"});
	let fontItalic = excel.addStyle({fontStyle: "italic"});
	let fontUnderlined = excel.addStyle({fontStyle: "underlined"});
	let fontColored = excel.addStyle({fontColor: "#ff3300"});
	let fontFamily = excel.addStyle({fontFamily: "Comic Sans MS"});

	excel.insert({column: 2, row: 2, value: "Big Text", style: fontBig});
	excel.insert({column: 2, row: 3, value: "Small Text", style: fontSmall});
	excel.insert({column: 2, row: 4, value: "Bold Text", style: fontBold});
	excel.insert({column: 2, row: 5, value: "Italic Text", style: fontItalic});
	excel.insert({column: 2, row: 6, value: "Underlined Text", style: fontUnderlined});
	excel.insert({column: 2, row: 7, value: "Colored Text", style: fontColored});
	excel.insert({column: 2, row: 8, value: "Font Family", style: fontFamily});

	// Create a complex style rule.
	let headerStyle = excel.addStyle({
		backgroundColor: "#a6a6a6",
		fontColor: "#ffffff",
		fontStyle: "bold",
		borderBottom: "medium"
	});

	// Apply a complex style on auto insert, and choosing an specific row.
	excel.insert(["Background Colors", "Font Styling", "Text Align", "Text Format", "Borders", "Conditional", "Custom"], {style: headerStyle, row: 1});

	// Insert inline styles.
	excel.insert({column: 3, row: 2, value: "Left Align", style: {textAlign: "left"}});
	excel.insert({column: 3, row: 3, value: "Center Align", style: {textAlign: "center"}});
	excel.insert({column: 3, row: 4, value: "Right Align", style: {textAlign: "right"}});

	// Align combined horizontally & vertically.
	excel.insert({column: 3, row: 5, value: "Top Align", style: {textAlign: "center top"}});
	excel.insert({column: 3, row: 6, value: "Middle Align", style: {textAlign: "center middle"}});
	excel.insert({column: 3, row: 7, value: "Bottom Align"});

	// Standard Formatting.
	excel.insert({column: 4, row: 2, value: 0.5, style: {format: "percentage"}});
	excel.insert({column: 4, row: 3, value: 500, style: {format: "currency"}});

	// Borders with Style.
	excel.insert({column: 5, row: 2, value: "All Borders", style: {border: "medium #ff3300"}});
	excel.insert({column: 5, row: 3, value: "Multiple Different", style: {borders: "medium #ff3300, thin #33cc33, dotted #ff3300, medium #ff3300"}});
	excel.insert({column: 5, row: 4, value: "Right Thin", style: {borderRight: "thin #33cc33"}});
	excel.insert({column: 5, row: 5, value: "Top Dotted", style: {borderTop: "dotted #ff3300"}});
	excel.insert({column: 5, row: 6, value: "Bottom Medium", style: {borderBottom: "medium #ff3300"}});

	// Freeze row/cols.
	excel.freeze({sheet: sheet, column: 1, row: 1});

	// Conditional Row Gradient Formatting.
	excel.setConditional({sheet: sheet, column: 6, start: 2});
	excel.insert({column: 6, row: 2, value: 0});
	excel.insert({column: 6, row: 3, value: 10});
	excel.insert({column: 6, row: 4, value: 100});
	excel.insert({column: 6, row: 5, value: 1000});
	excel.insert({column: 6, row: 6, value: 10000});

	// Condition Row Colors Types.
	excel.setConditional({sheet: sheet, column: 7, start: 2, type: "specific"});
	excel.insert({column: 7, row: 2, value: 0});
	excel.insert({column: 7, row: 3, value: 1});
	excel.insert({column: 7, row: 4, value: 0});
	excel.insert({column: 7, row: 5, value: 1});
	excel.insert({column: 7, row: 6, value: 1});

	excel.export("Excelliarmus-Styles");
};