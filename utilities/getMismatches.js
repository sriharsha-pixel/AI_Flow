
function getMismatches(prodData, testData) {
    const mismatches = [];

    prodData.forEach((prodRow, index) => {
        const testRow = testData[index];
        if (!testRow) {
            mismatches.push({
                Row: index + 1,
                Issue: "No corresponding row in test data"
            });
            return;
        }
        Object.keys(prodRow).forEach((key) => {
            //const prodValue = prodRow[key]?.toString().trim();
            //const testValue = testRow[key]?.toString().trim();

            const prodValue = (prodRow[key] ?? "").toString().trim();
            const testValue = (testRow[key] ?? "").toString().trim();

            if (prodValue !== testValue) {
                mismatches.push({
                    Row: index + 1,
                    Column: key,
                    ProdValue: prodValue,
                    TestValue: testValue
                });
            }
        });
    });

    return mismatches;
}

module.exports = { getMismatches };
