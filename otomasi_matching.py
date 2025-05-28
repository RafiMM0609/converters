import json

# File paths
user_terdampak_path = 'user_terdampak.json'
data_user_old_path = 'data_user_old.json'
output_path = 'matched_users.json'

# Load user terdampak
with open(user_terdampak_path, encoding='utf-8') as f:
    user_terdampak = json.load(f)

# Load data user old
with open(data_user_old_path, encoding='utf-8') as f:
    data_user_old = json.load(f)

# Ambil semua nama dari user terdampak (pastikan key 'user' dan 'name')
nama_terdampak = set()
for u in user_terdampak.get('user', []):
    name = u.get('name')
    if name and isinstance(name, str):
        nama_terdampak.add(name.strip().upper())

# Cari kecocokan di data_user_old (key 'nama ')
matched = []
for d in data_user_old:
    nama_old = d.get('nama ')
    if nama_old and isinstance(nama_old, str):
        if nama_old.strip().upper() in nama_terdampak:
            matched.append(
                d
            )

# Simpan hasil
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(matched, f, ensure_ascii=False, indent=2)

print(f"Selesai. {len(matched)} data cocok disimpan di {output_path}")
