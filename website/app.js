document.addEventListener("DOMContentLoaded", function () {

    const response =
        document.getElementById("response");

    const templateSelect =
        document.getElementById("templateSelect");

    const dataInput =
        document.getElementById("dataInput");

    const enterData =
        document.getElementById("enterData");

    const spreadsheetSearch =
        document.getElementById("spreadsheetSearch");

    const spreadsheetList =
        document.getElementById("spreadsheetList");

    const helpButton =
        document.getElementById("helpButton");

    const openTransactionSearch =
        document.getElementById("openTransactionSearch");

    const transactionCount =
        document.getElementById("transactionCount");

    function loadSpreadsheets(search = "") {

        fetch(
            `/api/spreadsheets?search=${encodeURIComponent(search)}`
        )
            .then(result => result.json())
            .then(data => {

                spreadsheetList.innerHTML = "";

                data.spreadsheets.forEach(filename => {

                    const item =
                        document.createElement("li");

                    item.textContent = filename;

                    item.addEventListener("click", function () {

                        window.currentSpreadsheet = filename;

                        response.textContent =
                            `Current spreadsheet: ${filename}`;

                        document.querySelectorAll(
                            "#spreadsheetList li"
                        ).forEach(li => {
                            li.classList.remove("selected");
                        });

                        item.classList.add("selected");
                    });

                    spreadsheetList.appendChild(item);
                });
            })
            .catch(error => {

                console.error(error);

                response.textContent =
                    "Could not load spreadsheets.";
            });
    }

    loadSpreadsheets();


    spreadsheetSearch.addEventListener("input", function () {

        loadSpreadsheets(
            spreadsheetSearch.value
        );
    });


    fetch("/api/templates")
        .then(result => result.json())
        .then(data => {

            templateSelect.innerHTML = "";

            data.templates.forEach(template => {

                const option =
                    document.createElement("option");

                option.value = template;
                option.textContent = template;

                templateSelect.appendChild(option);
            });
        })
        .catch(error => {

            console.error(error);

            response.textContent =
                "Could not load templates.";
        });


    enterData.addEventListener("click", function () {

        const text =
            dataInput.value.trim();

        const template =
            templateSelect.value;

        if (!text) {

            response.textContent =
                "Please enter some information.";

            return;
        }

        fetch("/api/transaction", {

            method: "POST",

            headers: {
                "Content-Type": "application/json"
            },

            body: JSON.stringify({
                text: text,
                template: template
            })
        })
            .then(result => result.json())
            .then(data => {

                if (data.error) {

                    response.textContent =
                        data.error;

                    return;
                }

                response.textContent =
                    "Transaction received by Phillip.";

                dataInput.value = "";

                dataInput.focus();

                loadSpreadsheets();
            })
            .catch(error => {

                console.error(error);

                response.textContent =
                    "Could not connect to Phillip.";
            });
    });


    document.getElementById("clearData")
        .addEventListener("click", function () {

            dataInput.value = "";

            dataInput.focus();

            response.textContent = "";
        });


    helpButton.addEventListener("click", function () {

        response.textContent =
            "Phillip Accounting Help: Select a spreadsheet template, enter accounting information, manage spreadsheets, search transactions, or use the kinematics calculators.";
    });


    openTransactionSearch.addEventListener("click", function () {

        const query = prompt(
            "Search transactions:"
        );

        if (query === null || !query.trim()) {
            return;
        }

        fetch(
            `/api/search-transactions?query=${encodeURIComponent(query.trim())}`
        )
            .then(result => result.json())
            .then(data => {

                if (data.error) {

                    response.textContent =
                        data.error;

                    return;
                }

                const transactionList =
                    document.getElementById("transactionList");

                transactionList.innerHTML = "";

                if (data.transactions.length === 0) {

                    response.textContent =
                        "No matching transactions found.";

                    transactionCount.textContent =
                        "0 transactions";

                    return;
                }

                data.transactions.forEach(transaction => {

                    const item =
                        document.createElement("li");

                    item.textContent =
                        Object.entries(transaction.data)
                            .map(([key, value]) =>
                                `${key}: ${value}`
                            )
                            .join(" | ");

                    transactionList.appendChild(item);
                });

                transactionCount.textContent =
                    `${data.transactions.length} transaction(s) found`;

                response.textContent =
                    "Transaction search complete.";
            })
            .catch(error => {

                console.error(error);

                response.textContent =
                    "Could not search transactions.";
            });
    });


    document.getElementById("searchTransactionsButton")
        .addEventListener("click", function () {

	    const query = prompt("Search transactions:")

	    if (query === null || !query.trim()) {
		return;
	    }

            openTransactionSearch.click();
        });


    document.getElementById("clearSpreadsheetSearch")
        .addEventListener("click", function () {

            spreadsheetSearch.value = "";

            loadSpreadsheets();
        });


    document.getElementById("clearNewSpreadsheet")
        .addEventListener("click", function () {

            document.getElementById(
                "newSpreadsheetName"
            ).value = "";
        });


    document.getElementById("selectSpreadsheet")
        .addEventListener("click", function () {

            const selected =
                document.querySelector(
                    "#spreadsheetList li.selected"
                );

            if (!selected) {

                response.textContent =
                    "Please select a spreadsheet first.";

                return;
            }

            const filename =
                selected.textContent;

            fetch("/api/select-spreadsheet", {

                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    filename: filename
                })
            })
                .then(result => result.json())
                .then(data => {

                    if (data.error) {

                        response.textContent =
                            data.error;

                        return;
                    }

                    window.currentSpreadsheet =
                        data.filename;

                    document.getElementById(
                        "currentSpreadsheet"
                    ).textContent =
                        `Current spreadsheet: ${data.filename}`;

                    response.textContent =
                        `Selected ${data.filename}.`;
                })
                .catch(error => {

                    console.error(error);

                    response.textContent =
                        "Could not select spreadsheet.";
                });
        });


    document.getElementById("deleteSpreadsheet")
        .addEventListener("click", function () {

            const selected =
                document.querySelector(
                    "#spreadsheetList li.selected"
                );

            if (!selected) {

                response.textContent =
                    "Please select a spreadsheet first.";

                return;
            }

            const filename =
                selected.textContent;

            if (filename === "Welcome.xlsx") {

                response.textContent =
                    "Welcome.xlsx cannot be deleted.";

                return;
            }

            if (!confirm(`Delete ${filename}?`)) {
                return;
            }

            fetch("/api/delete-spreadsheet", {

                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    filename: filename
                })
            })
                .then(result => result.json())
                .then(data => {

                    if (data.error) {

                        response.textContent =
                            data.error;

                        return;
                    }

                    response.textContent =
                        `${data.filename} deleted.`;

                    loadSpreadsheets();

                    document.getElementById(
                        "currentSpreadsheet"
                    ).textContent =
                        "Current spreadsheet: Welcome.xlsx";
                })
                .catch(error => {

                    console.error(error);

                    response.textContent =
                        "Could not delete spreadsheet.";
                });
        });


    document.getElementById("createSpreadsheet")
        .addEventListener("click", function () {

            const input =
                document.getElementById(
                    "newSpreadsheetName"
                );

            const filename =
                input.value.trim();

            if (!filename) {

                response.textContent =
                    "Please enter a spreadsheet name.";

                return;
            }

            fetch("/api/create-spreadsheet", {

                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    filename: filename
                })
            })
                .then(result => result.json())
                .then(data => {

                    if (data.error) {

                        response.textContent =
                            data.error;

                        return;
                    }

                    response.textContent =
                        `${data.filename} created.`;

                    input.value = "";

                    loadSpreadsheets();
                })
                .catch(error => {

                    console.error(error);

                    response.textContent =
                        "Could not create spreadsheet.";
                });
        });


    document.getElementById("calculate1D")
        .addEventListener("click", function () {

            const data = {

                initial_position:
                    document.getElementById(
                        "initialPosition"
                    ).value,

                end_position:
                    document.getElementById(
                        "endPosition"
                    ).value,

                initial_velocity:
                    document.getElementById(
                        "initialVelocity"
                    ).value,

                end_velocity:
                    document.getElementById(
                        "endVelocity"
                    ).value,

                acceleration:
                    document.getElementById(
                        "acceleration"
                    ).value,

                time:
                    document.getElementById(
                        "time"
                    ).value
            };

            fetch("/api/kinematics/1d", {

                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify(data)
            })
                .then(result => result.json())
                .then(data => {

                    const result =
                        document.getElementById(
                            "kinematicsResult"
                        );

                    if (data.error) {

                        result.textContent =
                            data.error;

                        return;
                    }

                    result.textContent =
                        JSON.stringify(
                            data.result,
                            null,
                            2
                        );
                })
                .catch(error => {

                    console.error(error);

                    document.getElementById(
                        "kinematicsResult"
                    ).textContent =
                        "Could not calculate kinematics.";
                });
        });


    document.getElementById("calculate2D")
        .addEventListener("click", function () {

            const speed =
                parseFloat(
                    document.getElementById(
                        "initialVelocity2D"
                    ).value
                );

            const angle =
                parseFloat(
                    document.getElementById(
                        "angle2D"
                    ).value
                );

            const angleRadians =
                angle * Math.PI / 180;

            const data = {

                initial_height:
                    document.getElementById(
                        "initialHeight2D"
                    ).value,

                initial_position:
                    document.getElementById(
                        "initialPosition2D"
                    ).value,

                vertical_velocity:
                    speed * Math.sin(angleRadians),

                horizontal_velocity:
                    speed * Math.cos(angleRadians),

                angle: angle,

                speed: speed,

                planet:
                    document.getElementById(
                        "planet2D"
                    ).value
            };

            fetch("/api/kinematics/2d", {

                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify(data)
            })
                .then(result => result.json())
                .then(data => {

                    const result =
                        document.getElementById(
                            "kinematicsResult2D"
                        );

                    if (data.error) {

                        result.textContent =
                            data.error;

                        return;
                    }

                    result.textContent =
                        JSON.stringify(
                            data.result,
                            null,
                            2
                        );
                })
                .catch(error => {

                    console.error(error);

                    document.getElementById(
                        "kinematicsResult2D"
                    ).textContent =
                        "Could not calculate 2D kinematics.";
                });
        });

});
