"متابعة اوامر غيار 2025"
if (sheetName === 'متابعة اوامر غيار 2025') {
const tableRows = table.querySelectorAll('tbody tr');
const TARGET_COLUMNS =[8, 9, 11, 13, 16, 18, 21];

tableRows.forEach(row = > {
const columns = row.querySelectorAll('td');
TARGET_COLUMNS.forEach(index = > {
if (index < columns.length) {
const thisCell = columns[index];
const prevCell = columns[index - 1];
if (thisCell & & prevCell) {
const currentDate = parseDate(thisCell.innerText);
const previousDate = parseDate(prevCell.innerText);
if (currentDate & & previousDate) {
if (currentDate > previousDate) {
thisCell.style.backgroundColor = COLORS.orange;
} else if (currentDate < previousDate) {
thisCell.style.backgroundColor = COLORS.fireRed;
} else {
thisCell.style.backgroundColor = COLORS.blue;
}
}
}
}
});
});
}