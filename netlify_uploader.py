import os
import requests
import zipfile
import io

def deploy_site(folder_path, site_id, token):
    """
    Zips the folder and deploys it to Netlify.
    
    :param folder_path: Path to the folder to deploy (e.g. QR_Patients)
    :param site_id: Netlify Site ID (API ID)
    :param token: Netlify Personal Access Token
    :return: (success, message_or_url)
    """
    if not os.path.exists(folder_path):
        return False, "Folder not found"

    print(f"Preparing to deploy {folder_path} to Netlify...")

    # 1. Create Zip in Memory
    try:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Archive name should be relative to the root folder
                    # e.g. if folder is QR_Patients, and file is QR_Patients/sub/file.txt
                    # arcname should be sub/file.txt
                    arcname = os.path.relpath(file_path, folder_path)
                    zip_file.write(file_path, arcname)
        
        zip_data = zip_buffer.getvalue()
        print(f"Zipped {len(zip_data)} bytes.")
    except Exception as e:
        return False, f"Error zipping files: {e}"

    # 2. Upload to Netlify
    url = f"https://api.netlify.com/api/v1/sites/{site_id}/deploys"
    headers = {
        "Content-Type": "application/zip",
        "Authorization": f"Bearer {token}"
    }

    try:
        response = requests.post(url, headers=headers, data=zip_data)
        
        if response.status_code == 200:
            data = response.json()
            deploy_url = data.get('ssl_url') or data.get('url')
            # Netlify usually returns the deploy specific URL, but we want the main site URL usually.
            # However, for verification, the deploy URL is fine. 
            # Actually, let's return the main site URL if possible, or just the deploy URL.
            # The 'url' field in response is usually the deploy preview URL (e.g. 64b...--site.netlify.app)
            # The 'ssl_url' is the main custom domain or netlify subdomain.
            return True, f"Deployed Successfully! URL: {deploy_url}"
        else:
            return False, f"Netlify Error {response.status_code}: {response.text}"
    except Exception as e:
        return False, f"Upload Request Error: {e}"
