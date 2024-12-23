// Handle Excel data loading
async function loadExcelData() {
    try {
        console.log('Attempting to fetch product data...');
        const response = await fetch('/get_product_data');
        const data = await response.json();

        console.log('Received response:', data);

        if (!response.ok) {
            throw new Error(data.error || `HTTP error! status: ${response.status}`);
        }

        if (data.success && data.data) {
            console.log('Successfully retrieved data:', data.data);
            populateFormFields(data.data);
        } else {
            console.error('Error from server:', data.error);
            throw new Error(data.error || 'Unknown server error');
        }
    } catch (error) {
        console.error('Error loading data:', error);
        const productSelect = document.getElementById('product-select');
        const productDetails = document.getElementById('product-details');

        // Clear existing options
        productSelect.innerHTML = '<option value="">Unable to load screw press data</option>';
        productSelect.disabled = true;

        // Create error message element
        const errorDiv = document.createElement('div');
        errorDiv.className = 'error-message';
        errorDiv.textContent = error.message || 'Failed to load screw press data. Please ensure the Excel file is properly uploaded and try again.';

        // Remove any existing error messages
        const existingError = productDetails.parentElement.querySelector('.error-message');
        if (existingError) {
            existingError.remove();
        }

        // Add new error message
        productDetails.parentElement.insertBefore(errorDiv, productDetails);
    }
}

// Populate form fields with Excel data
function populateFormFields(data) {
    console.log('Starting to populate form fields with data:', data);
    const productSelect = document.getElementById('product-select');

    // Clear existing options and error messages
    productSelect.innerHTML = '<option value="">Select a screw press...</option>';
    const existingError = document.querySelector('.error-message');
    if (existingError) {
        existingError.remove();
    }

    if (!data.screw_presses || !Array.isArray(data.screw_presses)) {
        console.error('Invalid data format:', data);
        throw new Error('Invalid data format received from server');
    }

    let validProducts = 0;

    // Populate products dropdown
    data.screw_presses.forEach(product => {
        // Skip empty rows or the control boxes row
        if (!product['Item Name (MD 300 Series)'] || 
            product['Item Name (MD 300 Series)'].toLowerCase().includes('control boxes')) {
            return;
        }

        validProducts++;
        const option = document.createElement('option');
        option.value = product['GDS Part No'] || product['Item Name (MD 300 Series)'];
        option.textContent = product['Item Name (MD 300 Series)'];

        // Store all product details in dataset
        option.dataset.manufacturer = product['Manufacturer'] || '-';
        option.dataset.partNumber = product['Mivalt Part Number'] || '-';
        option.dataset.gdsPartNo = product['GDS Part No'] || '-';
        option.dataset.power = product['Power'] || '-';
        option.dataset.material = product['Material'] || '-';
        option.dataset.leadTime = product['Lead Time'] || '-';
        option.dataset.costEuro = (parseFloat(product['Cost (Euro)']) || 0).toFixed(2);
        option.dataset.costUsd = (parseFloat(product['Cost USD']) || 0).toFixed(2);

        productSelect.appendChild(option);
    });

    console.log(`Added ${validProducts} valid products to dropdown`);

    if (validProducts === 0) {
        const errorDiv = document.createElement('div');
        errorDiv.className = 'error-message';
        errorDiv.textContent = 'No valid screw press products found in the data.';
        productSelect.parentElement.insertBefore(errorDiv, productSelect.nextSibling);
        productSelect.disabled = true;
    } else {
        productSelect.disabled = false;
    }

    // Add change event to update details when product is selected
    productSelect.addEventListener('change', updateProductDetails);
}

// Update product details in the form
function updateProductDetails(event) {
    const selectedOption = event.target.options[event.target.selectedIndex];
    const updateElement = (id, value) => {
        const element = document.getElementById(id);
        if (element) {
            if (element.tagName === 'INPUT') {
                element.value = value;
            } else {
                element.textContent = value;
            }
        }
    };

    if (selectedOption.value) {
        // Update all fields with dataset values
        updateElement('manufacturer', selectedOption.dataset.manufacturer);
        updateElement('part-number', selectedOption.dataset.partNumber);
        updateElement('gds-part-no', selectedOption.dataset.gdsPartNo);
        updateElement('power', selectedOption.dataset.power);
        updateElement('material', selectedOption.dataset.material);
        updateElement('lead-time', selectedOption.dataset.leadTime);
        updateElement('cost-euro', selectedOption.dataset.costEuro);
        updateElement('unit-price', selectedOption.dataset.costUsd);
        calculateTotal();
    } else {
        // Reset fields if no option selected
        ['manufacturer', 'part-number', 'gds-part-no', 'power', 'material', 'lead-time'].forEach(id => {
            updateElement(id, '-');
        });
        updateElement('cost-euro', '');
        updateElement('unit-price', '');
        updateElement('total-amount', '');
    }
}

// Handle image preview
function handleImagePreview(input) {
    const preview = document.getElementById('logo-preview');
    if (!input.files || !input.files[0]) {
        preview.style.display = 'none';
        return;
    }

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        preview.src = e.target.result;
        preview.style.display = 'block';
    };

    reader.onerror = function(e) {
        console.error('Error reading file:', e);
        preview.style.display = 'none';
    };

    reader.readAsDataURL(file);
}

// Calculate invoice total
function calculateTotal() {
    const quantity = parseFloat(document.getElementById('quantity').value) || 0;
    const price = parseFloat(document.getElementById('unit-price').value) || 0;
    const total = quantity * price;
    document.getElementById('total-amount').value = total.toFixed(2);
}

// Initialize the page
document.addEventListener('DOMContentLoaded', function() {
    console.log('Page loaded, initializing...');
    loadExcelData();

    // Add event listeners
    const quantityInput = document.getElementById('quantity');
    const unitPriceInput = document.getElementById('unit-price');

    if (quantityInput) {
        quantityInput.addEventListener('input', calculateTotal);
    }
    if (unitPriceInput) {
        unitPriceInput.addEventListener('input', calculateTotal);
    }
});