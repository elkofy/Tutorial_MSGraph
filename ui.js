async function displayUI() {
    await signIn();

    // Display info from user profile
    const user = await getUser();
    const Tablehead = await getTableHead();
    const Tablebody = await getTableBody();

    let userName = document.getElementById('userName');
    userName.innerText = user.displayName;

    // Hide login button and initial UI
    let signInButton = document.getElementById('signin');
    signInButton.style = "display: none";
    let content = document.getElementById('content');
    content.style = "display: block";

    let SynchroniseButton = document.getElementById('synchronise');
    SynchroniseButton.style = "display: block";
    
    let addvalueButton = document.getElementById('addvalue');
    addvalueButton.style = "display: block";
    let table = document.getElementById('table');
    table.style = "display: block";
   
    for (let i = 0; i < Tablehead.length; i++) {
        let th = document.createElement('th');
        th.innerText = Tablehead[i];
        table.appendChild(th);
    }

    for (let i = 0; i < Tablebody.length; i++) {
        let tr = document.createElement('tr');
        for (let j = 0; j < Tablebody[i].length; j++) {
            let td = document.createElement('td');
            // convert the date to the correct format
            if (j == 1) {
                //check if it's a number
                if (isNaN(Tablebody[i][j])) {
                    td.innerText = Tablebody[i][j];
                } else {
                    td.innerText = excelDateToJSDate(Tablebody[i][j]);

                }

            } else {
                td.innerText = Tablebody[i][j];
            }
            tr.appendChild(td);
        }
        table.appendChild(tr);
    }

}


function excelDateToJSDate(excelDate) {
    let date = new Date(Math.round((excelDate - (25567 + 1)) * 86400 * 1000));
    let converted_date = date.toISOString().split('T')[0];
    return converted_date;
}

async function synchronise() {
    const PurchaseOrders = await getPurchaseOrders();
    let table = document.getElementById('table');

    let bodys = [];

    PurchaseOrders.forEach(async element => {
        bodys.push(Object.values(element));
    });
    let rows = {
        values: bodys
    }
    await postTableRow(rows).then(
        () => {
            for (let i = 0; i < PurchaseOrders.length; i++) {
                let tr = document.createElement('tr');
                let keys = Object.keys(PurchaseOrders[i]);
                for (let j = 0; j < keys.length; j++) {
                    let td = document.createElement('td');
                    if (j == 1) {
                        td.innerText = PurchaseOrders[i][keys[j]].slice(0, 10)
                    } else {
                        td.innerText = PurchaseOrders[i][keys[j]];
                    }

                    tr.appendChild(td);
                }
                table.appendChild(tr);
            }
        }
    )
}

async function Addtestvalue() {
    let testrow = {
        values: [
            ['test', '2022-02-09', 45.40]
        ]
    }
    await postTableRow(testrow).then(() => {
        let tr = document.createElement('tr');
        testrow.values.forEach(element => {
            console.log(element);
            element.forEach(e => {
                let td = document.createElement('td');
                td.innerText = e;
                tr.appendChild(td);
            })
        }
        );
        let table = document.getElementById('table');
        table.appendChild(tr);

    });

}

