// Handle Excel data loading
async function loadExcelData() {
    try {
        const response = await fetch('/get_product_data');
        const data = await response.json();

        if (data.success) {
            populateFormFields(data.data);
        } else {
            // Display error in a user-friendly way
            const productSelect = document.getElementById('product-select');
            const productDetails = document.getElementById('product-details');

            productSelect.innerHTML = '<option value="">Unable to load screw press data</option>';
            productSelect.disabled = true;

            // Show error message to user
            const errorDiv = document.createElement('div');
            errorDiv.className = 'error-message';
            errorDiv.textContent = data.error || 'Error loading Excel data. Please ensure the Excel file is properly configured and try again.';
            productDetails.parentElement.insertBefore(errorDiv, productDetails);

            console.error('Error loading Excel data:', data.error);
        }
    } catch (error) {
        console.error('Error:', error);
    }
}

// Populate form fields with Excel data
function populateFormFields(data) {
    const productSelect = document.getElementById('product-select');

    // Clear existing options
    productSelect.innerHTML = '<option value="">Select a screw press...</option>';

    // Populate products dropdown
    data.screw_presses.forEach(product => {
        const option = document.createElement('option');
        option.value = product['Mivalt Part Number'];
        option.textContent = `${product['Item Name (MD 300 Series)']}`;
        // Store all product details in dataset
        option.dataset.manufacturer = product['Manufacturer'];
        option.dataset.partNumber = product['Mivalt Part Number'];
        option.dataset.gdsPartNo = product['GDS Part No'];
        option.dataset.power = product['Power'];
        option.dataset.material = product['Material'];
        option.dataset.leadTime = product['Lead Time'];
        option.dataset.costUsd = product['Cost USD'];
        productSelect.appendChild(option);
    });

    // Add change event to update details when product is selected
    productSelect.addEventListener('change', function() {
        const selectedOption = this.options[this.selectedIndex];
        updateProductDetails(selectedOption);
    });
}

// Update product details in the form
function updateProductDetails(selectedOption) {
    if (selectedOption.value) {
        document.getElementById('manufacturer').textContent = selectedOption.dataset.manufacturer || '-';
        document.getElementById('part-number').textContent = selectedOption.dataset.partNumber || '-';
        document.getElementById('gds-part-no').textContent = selectedOption.dataset.gdsPartNo || '-';
        document.getElementById('power').textContent = selectedOption.dataset.power || '-';
        document.getElementById('material').textContent = selectedOption.dataset.material || '-';
        document.getElementById('lead-time').textContent = selectedOption.dataset.leadTime || '-';
        document.getElementById('unit-price').value = selectedOption.dataset.costUsd || '';
        calculateTotal();
    } else {
        // Reset fields if no option selected
        ['manufacturer', 'part-number', 'gds-part-no', 'power', 'material', 'lead-time'].forEach(id => {
            document.getElementById(id).textContent = '-';
        });
        document.getElementById('unit-price').value = '';
        document.getElementById('total-amount').value = '';
    }
}

// Handle image preview
function handleImagePreview(input) {
    const preview = document.getElementById('logo-preview');
    const file = input.files[0];

    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            preview.src = e.target.result;
            preview.style.display = 'block';
        };
        reader.readAsDataURL(file);
    }
}

// Calculate invoice total
function calculateTotal() {
    const quantity = document.getElementById('quantity').value;
    const price = document.getElementById('unit-price').value;
    const total = quantity * price;
    document.getElementById('total-amount').value = total.toFixed(2);
}

// Initialize the page
document.addEventListener('DOMContentLoaded', function() {
    loadExcelData();

    // Add event listeners
    document.getElementById('quantity').addEventListener('input', calculateTotal);
    document.getElementById('unit-price').addEventListener('input', calculateTotal);
});