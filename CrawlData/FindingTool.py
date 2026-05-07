import csv

def find_person(csv_file, name, born_year):
    results = []
    with open(csv_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get('Name') and row.get('Birthday'):
                birthday_year = row['Birthday'].split('/')[-1]
                if name.lower() in row['Name'].lower() and birthday_year == str(born_year):
                    results.append(row)
    return results

if __name__ == "__main__":
    csv_path = "NN020.csv"
    search_name = input("Enter name to search: ")
    search_year = input("Enter born year to search: ")
    matches = find_person(csv_path, search_name, search_year)
    if matches:
        print("Found entries:")
        for match in matches:
            print(match)
    else:
        print("No matching entries found.")