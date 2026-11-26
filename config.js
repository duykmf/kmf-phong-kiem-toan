// ===== CẤU HÌNH – anh chỉ sửa file này =====
window.KMF_CONFIG = {
  // Mã ứng dụng (Application/client ID) lấy ở Entra > App registrations > KMF Phong Kiem Toan
  clientId: 'ac71f956-8879-4005-a795-a2f08e6179cd',
  // Mã tổ chức (Directory/tenant ID) của KMF
  tenantId: 'ab5d39de-9d6d-4c03-8619-eff2db7a3f98',
  // Trang SharePoint chứa dữ liệu (List được tạo tự động ở đây)
  siteHost: 'kmfaudit.sharepoint.com',
  sitePath: '/sites/Migration',
  // Email những người được bấm "Thiết lập" và "Nạp dự án"
  admins: ['duy.vk@kmf-audit.com'],
  // Số giây giữa 2 lần làm mới dữ liệu
  refreshSeconds: 8
};
