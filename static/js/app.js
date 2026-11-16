document.addEventListener('DOMContentLoaded', () => {
  const toggle = document.getElementById('menuToggle');
  const sidebar = document.getElementById('sidebar');
  if (toggle && sidebar) toggle.addEventListener('click', () => sidebar.classList.toggle('open'));
  const fileInput = document.getElementById('fileInput');
  const fileName = document.getElementById('fileName');
  if (fileInput && fileName) fileInput.addEventListener('change', () => {
    fileName.textContent = fileInput.files && fileInput.files.length ? fileInput.files[0].name : 'Pilih file dari komputer';
  });
});
