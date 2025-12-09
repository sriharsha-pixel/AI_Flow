const { faker } = require("@faker-js/faker");

exports.FakerDataPage = class FakerDataPage {
    fakerData() {
        const data = {
            firstname: faker.name.firstName(),
            lastname: faker.name.lastName(),
            totalprice: faker.number.int({ min: 100, max: 1000 }), 
            depositpaid: faker.datatype.boolean(),
            bookingdates: {
                checkin: faker.date.past().toISOString(),
                checkout: faker.date.future().toISOString(),
            },
            additionalneeds: faker.lorem.words()
        };
        return data;
    }
}
