// Excelliarmus - Your tiny excel generator.
// Michele LEO 2018
// github.com/leodau

let button = document.getElementsByTagName("button")[0];

button.onclick = (event) => {
	let excel = new Excelliarmus();

	// Create a sheet.
	excel.addSheet();

	// Insert at given col/row.
	excel.insert({column: 1, row: 1, value: "This"});
	excel.insert({column: 2, row: 2, value: "is"});
	excel.insert({column: 3, row: 3, value: "so"});
	excel.insert({column: 4, row: 4, value: "easy"});

	// Export your file!
	excel.export();
};