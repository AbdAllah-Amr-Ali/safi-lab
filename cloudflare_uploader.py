import os
import requests
import json
import hashlib

def calculate_file_hash(filepath):
    """Calculates the SHA256 hash of a file."""
    sha256_hash = hashlib.sha256()
    with open(filepath, "rb") as f:
        # Read and update hash string value in blocks of 4K
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def upload_files(file_paths, project_name, account_id, api_token):
    """
    Uploads a list of files to Cloudflare Pages.
    
    :param file_paths: Dictionary { remote_path: local_path }
    :param project_name: Name of the Cloudflare Pages project.
    :param account_id: Cloudflare Account ID.
    :param api_token: Cloudflare API Token.
    :return: (success, message)
    """
    url = f"https://api.cloudflare.com/client/v4/accounts/{account_id}/pages/projects/{project_name}/deployments"
    
    headers = {
        "Authorization": f"Bearer {api_token}"
    }
    
    try:
        files_to_close = []
        payload_files = []
        manifest = {}
        
        for remote_path, local_path in file_paths.items():
            if os.path.exists(local_path):
                # Calculate Hash
                file_hash = calculate_file_hash(local_path)
                
                # Add to manifest
                # Cloudflare expects path starting with /
                manifest_path = "/" + remote_path.lstrip("/")
                manifest[manifest_path] = file_hash
                
                # Prepare file for upload
                f = open(local_path, 'rb')
                files_to_close.append(f)
                payload_files.append(('files', (local_path, f)))
            else:
                print(f"File not found: {local_path}")

        if not payload_files:
            return False, "No files found to upload."

        print(f"Uploading {len(payload_files)} files to Cloudflare Pages ({project_name})...")
        
        # Add manifest to payload
        # For Direct Upload, the manifest is a JSON string in the 'manifest' form field
        data = {"manifest": json.dumps(manifest)}
        
        response = requests.post(url, headers=headers, files=payload_files, data=data)
        
        # Close files
        for f in files_to_close:
            f.close()
            
        if response.status_code == 200:
            data = response.json()
            if data.get('success'):
                deployment_url = data['result']['url']
                return True, f"Deployed successfully! URL: {deployment_url}"
            else:
                return False, f"Upload failed: {data['errors'][0]['message']}"
        else:
            return False, f"HTTP Error {response.status_code}: {response.text}"
            
    except Exception as e:
        return False, f"Exception during upload: {str(e)}"
