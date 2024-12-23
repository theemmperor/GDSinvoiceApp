// Handle Excel data loading
async function loadExcelData() {
    try {
        const response = await fetch('/get_excel_data');
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
    const customerSelect = document.getElementById('customer-select');
    
    // Populate products dropdown
    data.products.forEach(product => {
        const option = document.createElement('option');
        option.value = product.id;
        option.textContent = `${product.name} - $${product.price}`;
        productSelect.appendChild(option);
    });
    
    // Populate customers dropdown
    data.customers.forEach(customer => {
        const option = document.createElement('option');
        option.value = customer.id;
        option.textContent = customer.name;
        customerSelect.appendChild(option);
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
