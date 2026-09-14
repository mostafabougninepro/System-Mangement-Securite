def get_agent_photo(matricule):
    if not matricule or not str(matricule).strip():
        return None, "Matricule vide"
    
    target = str(matricule).strip().lower()
    
    # 1. البحث أولاً داخل ملف الـ ZIP مباشرة (أضمن وأسرع طريقة)
    if os.path.exists(ZIP_PATH):
        try:
            with zipfile.ZipFile(ZIP_PATH, 'r') as zip_ref:
                for file_name in zip_ref.namelist():
                    base_name = os.path.basename(file_name)
                    if not base_name: # تخطي المجلدات الفارغة
                        continue
                    name_part, _ = os.path.splitext(base_name)
                    if name_part.strip().lower() == target:
                        # استخراج الصورة المؤقتة فوراً وعرضها
                        extracted_path = zip_ref.extract(file_name, path=os.path.join(BASE_DIR, "_temp_extracted"))
                        return extracted_path, "Photo trouvée dans ZIP"
        except Exception:
            pass

    # 2. البحث الاحتياطي داخل المجلدات المفكوكة
    if os.path.exists(PHOTOS_DIR):
        try:
            for root, dirs, files in os.walk(PHOTOS_DIR):
                for file_name in files:
                    name_part, _ = os.path.splitext(file_name)
                    if name_part.strip().lower() == target:
                        return os.path.join(root, file_name), "Photo trouvée dans Dossier"
        except Exception:
            pass
            
    return None, "Photo non trouvable"
