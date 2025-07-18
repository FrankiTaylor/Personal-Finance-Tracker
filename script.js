document.addEventListener('DOMContentLoaded', function() {
    // Ensure categories are populated on DOMContentLoaded for setup.html
    if (document.getElementById('categorySelect')) {
        populateCategories();
    }

    if(document.getElementById('accountSelect')) {
        populateAccounts();
    }

    // // On index.html, populate the dropdown with custom categories from localStorage
    // var categorySelect = document.getElementById('categorySelect');
    // if (categorySelect) {
    //     let categories = JSON.parse(localStorage.getItem('customCategories') || '[]');
    //     categories.forEach(function(cat) {
    //         // Avoid duplicates
    //         if (![...categorySelect.options].some(opt => opt.value === cat)) {
    //             var newOption = document.createElement('option');
    //             newOption.value = cat;
    //             newOption.text = cat;
    //             categorySelect.add(newOption);
    //         }
    //     });
    // }
});

const defaultCategories = [
    'Bank Checking',
    'Savings',
    'Credit Card'
];

const defaultAccounts = ['Setup account'];

function getAllAccounts() {
    let accounts = JSON.parse(localStorage.getItem('accounts') || '[]');
    let defaults = JSON.parse(localStorage.getItem('defaultAccounts') || '[]');

    // Add default account only if no custom accounts are present
    if(!defaults.length) {
        localStorage.setItem('defaultAccounts', JSON.stringify(defaultAccounts));
        defaults = [...defaultAccounts];
        let allAccounts = [...new Set([...defaults, ...accounts])];
        return allAccounts;
    }
    // Add all the added accounts excluding the default value
    else{
        let allAccounts = [...new Set([...accounts])];
        return allAccounts;
    }
}

function populateAccounts() {
    var accountSelect = document.getElementById('accountSelect');

    if (accountSelect) {
        accountSelect.innerHTML = '';
        
        let allAccounts = getAllAccounts();

        // Add default account only if no custom accounts are present
        if(!allAccounts.length){
            var defaultOption = document.createElement('option');
            defaultOption.value = defaultAccounts;
            defaultOption.text = defaultAccounts;
            accountSelect.add(defaultOption);
        } 

        allAccounts.forEach(function(acc) {
            var newOption = document.createElement('option');
            newOption.value = acc;
            newOption.text = acc;
            accountSelect.add(newOption);
        });
    }
}

function getAllCategories() {
    let custom = JSON.parse(localStorage.getItem('customCategories') || '[]');
    let defaults = JSON.parse(localStorage.getItem('defaultCategories') || '[]');

    // Ensure defaults are set if not already in localStorage
    if (!defaults.length) {
        localStorage.setItem('defaultCategories', JSON.stringify(defaultCategories));
        defaults = [...defaultCategories];
    }

    // Combine and remove duplicates
    let allCategories = [...new Set([...defaults, ...custom])];

    return allCategories;
}

function populateCategories() {
    var categorySelect = document.getElementById('categorySelect');
    if (categorySelect) {
        categorySelect.innerHTML = ''; // Clear existing options

        let allCategories = getAllCategories();
        allCategories.forEach(function(cat) {
            var newOption = document.createElement('option');
            newOption.value = cat;
            newOption.text = cat;
            categorySelect.add(newOption);
        });
    }
}

function deleteCategory() {
    var categorySelect = document.getElementById('categorySelect');
    if (categorySelect && categorySelect.selectedIndex !== -1) {
        let toDelete = categorySelect.value;
        let custom = JSON.parse(localStorage.getItem('customCategories') || '[]');
        let defaults = JSON.parse(localStorage.getItem('defaultCategories') || '[]');

        // Try to remove from custom first
        if (custom.includes(toDelete)) {
            custom = custom.filter(cat => cat !== toDelete);
            localStorage.setItem('customCategories', JSON.stringify(custom));
        } else if (defaults.includes(toDelete)) {
            defaults = defaults.filter(cat => cat !== toDelete);
            localStorage.setItem('defaultCategories', JSON.stringify(defaults));
        }

        populateCategories();
    }
}

// Expose for inline HTML usage
window.deleteCategory = deleteCategory;

document.getElementById('excelFileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Convert 'Date' column values to string if present
            if (jsonData.length > 0) {
                const dateColIndex = jsonData[0].findIndex(h => h.toLowerCase() === 'date');
                if (dateColIndex !== -1) {
                    for (let i = 1; i < jsonData.length; i++) {
                        let cell = jsonData[i][dateColIndex];
                        if (typeof cell === 'number') {
                            // Convert Excel date serial to string (YYYY-MM-DD)
                            const date = XLSX.SSF.parse_date_code(cell);
                            if (date) {
                                const yyyy = date.y;
                                const mm = String(date.m).padStart(2, '0');
                                const dd = String(date.d).padStart(2, '0');
                                jsonData[i][dateColIndex] = `${yyyy}-${mm}-${dd}`;
                            } else {
                                jsonData[i][dateColIndex] = String(cell);
                            }
                        } else if (cell !== undefined && cell !== null) {
                            jsonData[i][dateColIndex] = String(cell);
                        }
                    }
                }
            }

            displayData(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }
});

function displayData(data) {
    const dataDisplay = document.getElementById('dataDisplay');
    dataDisplay.innerHTML = '';

    if (data.length === 0) {
        dataDisplay.innerText = 'No data found in the file.';
        return;
    }

    // Create a table to display the data
    const table = document.createElement('table');
    table.classList.add('data-table');

    // Create header row
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
        const th = document.createElement('th');
        th.innerText = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create rows for each data entry
    data.slice(1).forEach(row => {
        const tr = document.createElement('tr');
        row.forEach(cell => {
            const td = document.createElement('td');
            td.contentEditable = 'true'; // Make cells editable
            td.innerText = cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    dataDisplay.appendChild(table);
}

function addAccount() {
    var categorySelectInput = document.getElementById('categorySelect');
    var categorySelectValue = categorySelectInput.value;
    var newNameValueInput = document.getElementById('inputName');
    var newNameValue = newNameValueInput.value.trim();

    if (newNameValue !== '') {
        // Save to localStorage
        newNameValue += ' - ' + categorySelectValue;
        let accounts = JSON.parse(localStorage.getItem('accounts') || '[]');
        if (!accounts.includes(newNameValue)) {
            accounts.push(newNameValue);
            localStorage.setItem('accounts', JSON.stringify(accounts));
            newNameValueInput.value = '';

            alert('Account added! It will appear in the dropdown on the main page.');

            // Populate categories to update the dropdown
            populateAccounts();
        } else {
            alert('Account already exists!');
        }
    }
}

// Added for testing functionality
function clearLocalStorage() {
    localStorage.clear();
    populateCategories();
    alert('LocalStorage cleared!');
}