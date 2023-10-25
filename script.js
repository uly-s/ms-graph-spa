const fs = require('fs');

var data = [];

fs.readFile('data.json', 'utf8', (err, file) => {
    if (err) {
        console.error(err);
        return;
    }
    data = JSON.parse(file);
    console.log(data.length);
    dedupe();
    console.log(data.length);
});

function compare(contact1, contact2) {
    if (contact1.displayName === contact2.displayName) {
    return true;
    }
    return false;
}

function dedupe() {
    for (let i = 0; i < data.length; i++) {
        for (let j = i+1; j < data.length; j++) {
            if (compare(data[i], data[j])) {
                data.pop(j);
            }
        }
    }

}




