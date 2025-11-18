document.addEventListener('DOMContentLoaded', function() {
    
    console.log('Demo.js loaded successfully!');
    
    // --- 1. Xử lý logic đăng nhập giả lập ---
    // --- Xử lý logic đăng ký giả lập ---
const signUpForm = document.getElementById('signUpForm');

if (signUpForm) {
    signUpForm.addEventListener('submit', function(e) {
        e.preventDefault(); // Ngăn form gửi đi

        const username = document.getElementById('newUsernameInput').value;
        const email = document.getElementById('emailInput').value;
        const password = document.getElementById('newPasswordInput').value;
        const confirmPassword = document.getElementById('confirmPasswordInput').value;

        // 1. Kiểm tra xác nhận mật khẩu
        if (password !== confirmPassword) {
            alert('Lỗi: Mật khẩu và Xác nhận Mật khẩu không khớp!');
            document.getElementById('newPasswordInput').value = '';
            document.getElementById('confirmPasswordInput').value = '';
            return;
        }
        
        // 2. Kiểm tra độ dài mật khẩu tối thiểu (tùy chọn)
        if (password.length < 3) {
            alert('Lỗi: Mật khẩu phải có ít nhất 3 ký tự.');
            return;
        }

        // --- ĐĂNG KÝ THÀNH CÔNG GIẢ LẬP ---
        alert(`Đăng ký thành công cho tài khoản "${username}"! Bạn đã được đăng nhập.`);
        
        // Ẩn modal Đăng ký
        const signUpModal = bootstrap.Modal.getInstance(document.getElementById('signUpModal'));
        if (signUpModal) {
            signUpModal.hide();
        }

        // ***** ĐOẠN ĐÃ CẬP NHẬT: Tự động đăng nhập và hiển thị tên *****
        updateHeaderAfterSignIn(username);
        // *************************************************************

        // Xóa nội dung các trường sau khi đăng ký
        signUpForm.reset();
    });
}
// ... (Đặt trong document.addEventListener('DOMContentLoaded', function() { ... });)

    // --- Xử lý logic đăng nhập giả lập ---
    const signInForm = document.getElementById('signInForm');

    if (signInForm) {
        signInForm.addEventListener('submit', function(e) {
            e.preventDefault(); // Ngăn form gửi đi

            const username = document.getElementById('usernameInput').value;
            const password = document.getElementById('passwordInput').value;
            
            // TÀI KHOẢN MẪU:
            const VALID_USERNAME = 'user';
            const VALID_PASSWORD = '123'; 

            // Kiểm tra thông tin đăng nhập
            if (username === VALID_USERNAME && password === VALID_PASSWORD) {
                // Đăng nhập thành công
                alert(`Xin chào, ${username}! Bạn đã đăng nhập thành công.`);
                
                // Ẩn modal sau khi đăng nhập thành công
                const signInModal = bootstrap.Modal.getInstance(document.getElementById('signInModal'));
                if (signInModal) {
                    signInModal.hide();
                }
                
                // ***** ĐOẠN QUAN TRỌNG: GỌI HÀM CẬP NHẬT HEADER *****
                updateHeaderAfterSignIn(username); 
                // *************************************************

            } else {
                // Đăng nhập thất bại
                alert('Tên đăng nhập hoặc mật khẩu không đúng. Vui lòng thử lại.');
            }

            // Xóa trường mật khẩu để bảo mật
            document.getElementById('passwordInput').value = '';
        });
    }

// ...
    // --- 2. Hàm cập nhật Header sau khi đăng nhập ---
    function updateHeaderAfterSignIn(username) {
        const signInButtonContainer = document.querySelector('.col.d-flex.align-items-center.justify-content-end.gap-3');
        const signInButton = signInButtonContainer.querySelector('.btn-warning');

        if (signInButton) {
            // Xóa nút 'Sign in' cũ
            signInButton.remove();

            // Tạo nút mới hiển thị tên người dùng
            const welcomeButton = document.createElement('button');
            welcomeButton.className = 'btn btn-warning';
            welcomeButton.textContent = `Chào, ${username}`;
            welcomeButton.style.cursor = 'default'; // Không cho nhấn vào

            // Tạo nút Đăng xuất
            const signOutButton = document.createElement('button');
            signOutButton.className = 'btn btn-outline-warning';
            signOutButton.textContent = 'Đăng xuất';
            signOutButton.onclick = function() {
                // Logic đăng xuất: tải lại trang hoặc xóa thông tin đăng nhập
                alert('Đã đăng xuất.');
                window.location.reload(); 
            };
            
            // Thêm các nút mới vào header
            signInButtonContainer.appendChild(welcomeButton);
            signInButtonContainer.appendChild(signOutButton);
        }
    }


    // --- 3. Chức năng More Menu đã có từ trước (nên giữ lại) ---
    const moreMenuBtn = document.querySelector('.more-menu-btn');
    if (moreMenuBtn) {
        moreMenuBtn.addEventListener('click', function() {
            alert('Đã nhấn nút More Menu! (Chức năng này cần được lập trình thêm)');
        });
    }

});
