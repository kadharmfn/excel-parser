const xlsx = require('node-xlsx');
const fs = require('fs');


const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${__dirname}/test.xlsx`));
// console.log("workSheetsFromBuffer", workSheetsFromBuffer[0].data);

const idsToMatch = [];
const columnsOtherOne = [];
const columnsOtherTwo = [];

workSheetsFromBuffer[0].data.slice(1).forEach(item => {
    console.log(item);
    const idToMatch = item[0];
    const columnOtherOne = item[1];
    const columnOtherTwo = item[2];

    idsToMatch.push(idToMatch);
    columnsOtherOne.push(columnOtherOne);
    columnsOtherTwo.push(columnOtherTwo);
});

console.log(idsToMatch, columnsOtherOne, columnsOtherTwo);


const column2Unmatched = [];
const column3Unmatched = [];


findUnmatchedBotId = (column, item) => {

    let itemToMatchArray = item.toString().split(',');

    const filtered = itemToMatchArray.filter(row => {
        let itemToMatch = row;
        if (itemToMatch.startsWith('CORPPVT\\') || itemToMatch.startsWith('CWWPVT\\') ) {
            itemToMatch = itemToMatch.substring(itemToMatch.indexOf('\\')+1);
        } else if (itemToMatch.startsWith('09')) {
            itemToMatch = Number(itemToMatch);
        }
    
        console.log("itemToMatch -->", column.indexOf(itemToMatch));
        if (column.indexOf(itemToMatch) === -1 && column.indexOf(itemToMatch.toLowerCase()) === -1 ) {
            console.log('not matched', itemToMatch);
            return true;
        }
    
        return false;
    });

    // console.log("filtered", filtered);

    return filtered;

}

idsToMatch.forEach(item => {
    if (item !== undefined) {

        const Unmatched = findUnmatchedBotId(columnsOtherOne, item);
        if (Unmatched.length >0) {
            column2Unmatched.push([...Unmatched]);
        }
    
        // if (matchBotId(columnsOtherTwo, item)) {
        //     column3Unmatched.push(item);
        // }
    }
   
})


console.log("Workn Tech Unmatched", column2Unmatched);
var buffer = xlsx.build([{name: "Workintech", data: column2Unmatched}]); 

fs.writeFileSync("unmatched.xlsx", buffer);

console.log("buffer", buffer);


