// Handle Excel data loading
async function loadExcelData() {
    try {
        const response = await fetch('/get_product_data');
        const data = await response.json();

        if (data.success) {
            populateFormFields(data.data);
        } else {
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
    productSelect.innerHTML = '<option value="">Select a product...</option>';

    // Populate products dropdown
    data.screw_presses.forEach(product => {
        const option = document.createElement('option');
        option.value = product['Mivalt Part Number'];
        option.textContent = `${product['Item Name (MD 300 Series)']} - ${product['Manufacturer']}`;
        option.dataset.price = product['Cost USD'];
        productSelect.appendChild(option);
    });

    // Add change event to update price when product is selected
    productSelect.addEventListener('change', function() {
        const selectedOption = this.options[this.selectedIndex];
        if (selectedOption.dataset.price) {
            document.getElementById('unit-price').value = selectedOption.dataset.price;
            calculateTotal();
        }
    });
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
    document.getElementById('quantity').addEventListener('change', calculateTotal);
    document.getElementById('unit-price').addEventListener('change', calculateTotal);
});