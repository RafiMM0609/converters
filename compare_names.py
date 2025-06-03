import json

# Read file1.json
with open('file1.json', 'r', encoding='utf-8') as f:
    file1_data = json.load(f)

# Read file2.json
with open('file2.json', 'r', encoding='utf-8') as f:
    file2_data = json.load(f)

# Extract names from both files
file1_key = list(file1_data.keys())[0]
file2_key = list(file2_data.keys())[0]

file1_names = set()
for item in file1_data[file1_key]:
    if 'name' in item:
        file1_names.add(item['name'].strip().upper())

file2_names = set()
for item in file2_data[file2_key]:
    if 'name' in item:
        file2_names.add(item['name'].strip().upper())

# Find common names (case-insensitive)
common_names = file1_names.intersection(file2_names)

print("=== PERBANDINGAN NAMA ANTARA FILE1.JSON DAN FILE2.JSON ===\n")
print(f"Total nama di file1.json: {len(file1_names)}")
print(f"Total nama di file2.json: {len(file2_names)}")
print(f"Nama yang sama: {len(common_names)}\n")

if common_names:
    print("DAFTAR NAMA YANG SAMA:")
    print("-" * 50)
    for i, name in enumerate(sorted(common_names), 1):
        print(f"{i:2d}. {name}")
else:
    print("Tidak ada nama yang sama antara kedua file.")

# Find names that exist in file1 but not in file2
only_in_file1 = file1_names - file2_names
print(f"\nNama yang hanya ada di file1.json: {len(only_in_file1)}")

# Find names that exist in file2 but not in file1
only_in_file2 = file2_names - file1_names
print(f"Nama yang hanya ada di file2.json: {len(only_in_file2)}")

# Show detailed comparison for verification
print("\n=== DETAIL NAMA YANG SAMA DENGAN ID ===")
print("-" * 60)

for item1 in file1_data[file1_key]:
    if 'name' in item1:
        name1 = item1['name'].strip().upper()
        if name1 in common_names:
            # Find corresponding item in file2
            for item2 in file2_data[file2_key]:
                if 'name' in item2 and item2['name'].strip().upper() == name1:
                    print(f"Nama: {item1['name']}")
                    print(f"  File1 - ID: {item1['id']}, Client ID: {item1['client_id']}")
                    print(f"  File2 - ID: {item2['id']}, Client ID: {item2['client_id']}")
                    print()
                    break
