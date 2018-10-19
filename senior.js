// Excelliarmus - Your tiny excel generator.
// Michele LEO 2018
// github.com/leodau

let buttonC = document.getElementsByTagName("button")[2];

buttonC.onclick = (event) => {
	let excel = new Excelliarmus();
	let sheet = excel.addSheet("Grocery List");
	let list = [
		{name: "Apple", count: 3, each: 0.5},
		{name: "Avocado", count: 1, each: 1.5},
		{name: "Mango", count: 2, each: 2.5},
		{name: "Eggs", count: 12, each: 0.2},
		{name: "Milk", count: 1, each: 2.5},
		{name: "Cereal", count: 1, each: 3},
	];

	excel.insert(["ITEM", "COUNT", "EACH", "PRICE"]);
	// Use a formula per item!
	list.forEach((item, i) => {
		excel.insert([item.name, item.count, item.each, "=B" + (i + 2) + "*C" + (i + 2)]);
	});
	excel.insert({column: 3, row: list.length + 3, value: "TOTAL"});

	// Create a dynamic formula with Excelliarmus intToColumn tool.
	excel.insert({
		column: 4,
		row: list.length + 3,
		value: "=SUM(" + Excelliarmus.intToColumn(3) + "2:" + Excelliarmus.intToColumn(3) + (list.length + 1) + ")",
	});

	excel.export("Excelliarmus-Senior");
	console.log(excel);
};