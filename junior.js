// Excelliarmus - Your tiny excel generator.
// Michele LEO 2018
// github.com/leodau

let buttonB = document.getElementsByTagName("button")[1];
	
buttonB.onclick = (event) => {
	let excel = new Excelliarmus();

	// Setup sheets with a name.
	let players = excel.addSheet("Players");
	let mobs = excel.addSheet("Mobs");

	// Insert values horizontaly as rows.
	excel.insert(["Id", "NAME", "LEVEL", "ROLE"]);
	excel.insert(["1", "Leo", 0, "Admin"]);
	excel.insert(["2", "John", 10, "Moderator"]);
	excel.insert(["3", "Eureka", 1000, "Player"]);
	excel.insert(["4", "", 20000]);
	excel.insert(["5", "Gabi", "30", "Player"]);

	// Insert values as rows in a given sheet id.
	excel.insert(["Id", "NAME", "LEVEL", "ELEMENT"], mobs);
	excel.insert(["1", "Blob", "2", "Earth"], mobs);
	excel.insert(["2", "FireBlob", "10"], mobs);
	excel.insert(["3", "Phoenix", "20", "Fire"], mobs);
	excel.insert(["4", "Frog", "7", "Plant"], mobs);
	excel.insert(["5", "Snake", "30", "Plant"], mobs);

	// Export your file with a name.
	excel.export("Excelliarmus-Junior");
};