document.addEventListener('DOMContentLoaded', function() {
    // Initialize Flatpickr for date inputs
    flatpickr('input[type="date"]', {
        dateFormat: 'Y-m-d',
        allowInput: false, // Prevent manual typing to enforce calendar selection
        maxDate: '2030-12-31', // Optional: Set a reasonable future date limit
        minDate: '2020-01-01'  // Optional: Set a reasonable past date limit
    });

    // Initialize Chervic signature canvas
    const chervicCanvas = document.getElementById('chervic_signature_canvas');
    const chervicCtx = chervicCanvas.getContext('2d');
    const chervicDataInput = document.getElementById('chervic_signature_data');
    
    // Initialize Customer signature canvas
    const customerCanvas = document.getElementById('customer_signature_canvas');
    const customerCtx = customerCanvas.getContext('2d');
    const customerDataInput = document.getElementById('customer_signature_data');
    
    let isDrawing = false;
    
    function setupCanvas(canvas, ctx, dataInput) {
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.strokeStyle = '#000';
        
        canvas.addEventListener('mousedown', startDrawing);
        canvas.addEventListener('mousemove', draw);
        canvas.addEventListener('mouseup', stopDrawing);
        canvas.addEventListener('mouseout', stopDrawing);
        canvas.addEventListener('touchstart', handleTouchStart);
        canvas.addEventListener('touchmove', handleTouchMove);
        canvas.addEventListener('touchend', stopDrawing);
        
        function startDrawing(e) {
            isDrawing = true;
            const rect = canvas.getBoundingClientRect();
            const x = (e.clientX || (e.touches && e.touches.length > 0 ? e.touches[0].clientX : 0)) - rect.left;
            const y = (e.clientY || (e.touches && e.touches.length > 0 ? e.touches[0].clientY : 0)) - rect.top;
            ctx.beginPath();
            ctx.moveTo(x, y);
        }
        
        function draw(e) {
            if (!isDrawing) return;
            e.preventDefault();
            const rect = canvas.getBoundingClientRect();
            const x = (e.clientX || (e.touches && e.touches.length > 0 ? e.touches[0].clientX : 0)) - rect.left;
            const y = (e.clientY || (e.touches && e.touches.length > 0 ? e.touches[0].clientY : 0)) - rect.top;
            ctx.lineTo(x, y);
            ctx.stroke();
        }
        
        function stopDrawing() {
            if (isDrawing) {
                isDrawing = false;
                dataInput.value = canvas.toDataURL('image/png');
            }
        }
        
        function handleTouchStart(e) {
            if (e.touches.length > 0) {
                e.preventDefault();
                startDrawing(e);
            }
        }
        
        function handleTouchMove(e) {
            if (e.touches.length > 0) {
                e.preventDefault();
                draw(e);
            }
        }
    }
    
    setupCanvas(chervicCanvas, chervicCtx, chervicDataInput);
    setupCanvas(customerCanvas, customerCtx, customerDataInput);
    
    // Clear canvas when file is uploaded
    document.getElementById('chervic_signature').addEventListener('change', function() {
        chervicCtx.clearRect(0, 0, chervicCanvas.width, chervicCanvas.height);
        chervicDataInput.value = '';
    });
    document.getElementById('customer_signature').addEventListener('change', function() {
        customerCtx.clearRect(0, 0, customerCanvas.width, customerCanvas.height);
        customerDataInput.value = '';
    });
    
    // Update hidden inputs before form submission
    document.getElementById('ndaForm').addEventListener('submit', function() {
        chervicDataInput.value = chervicCanvas.toDataURL('image/png');
        customerDataInput.value = customerCanvas.toDataURL('image/png');
    });
});