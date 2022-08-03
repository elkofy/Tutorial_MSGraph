
const ExcelPath = 'sites/<drive-number>/drive/root:/<pathtofile>/Graphtuto.xlsx:/workbook/worksheets/Sheet1/'
// Create an authentication provider

const authProvider = {
    getAccessToken: async () => {
        // Call getToken in auth.js
        return await getToken();
    }
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
//Get user info from Graph
async function getUser() {
    ensureScope('user.read');
    return await graphClient
        .api('/me')
        .select('id,displayName')
        .get();
}

async function getTableHead() {
    let headers = []
    // Loop in order to get the head of the table
    for (var i = 0; i < 3; i++) {
        let endpoint = ExcelPath + `cell(row=0,column=${i})`;
        let col = await graphClient
            .api(endpoint)
            .get()
            .then((res) => {
                return res.text.flat(); // Get the text of the cell
            }).catch((err) => {
                console.log(err);
            });
        headers.push(col);
    }
    return headers.flat(); // Return the array of the headers
}

async function getTableBody() {
    let endpoint = ExcelPath + 'tables/Table_1/rows';
    return await graphClient
        .api(endpoint)
        .get()
        .then((res) => {
            return res.value;
        }).catch((err) => {
            console.log(err);
        })
        .then(
            (res) => {
                let body = [];
                for (var i = 0; i < res.length; i++) {
                    body.push(res[i].values.flat());
                }
                return body;
            }
        )
        .catch((err) => {
            console.log(err);
        });
}

/*Retrieve purchase orders from the mockapi API*/
async function getPurchaseOrders() {
    return await fetch('https://youtpurchaseorders.com/purchaseorders')
        .then(response => response.json())
        .then(data => {
            // console.log(data);
            data.forEach(element => {
                delete element.id;
                console.log(element.Date)

                element.Date = element.Date.substring(0, 10);
            });
            return data;
        }).catch((err) => {
            console.log(err);
        }
        );
}

async function postTableRow(row) {
    let endpoint = ExcelPath + 'tables/Table_1/rows';
    return await graphClient
        .api(endpoint)
        .post(row)
        .then((res) => {
            return res;
        }).catch((err) => {
            console.log(err);
        }
        );
}