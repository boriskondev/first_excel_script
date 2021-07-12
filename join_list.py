data = input().strip();

processed_data = []

while data != "End":
    processed_data.append(data)
    data = input().strip();

print(len(processed_data))

output_file_csv = "result.csv"

with open(output_file_csv, "w") as file:
    while processed_data:
        file.write(", ".join(processed_data[0:49]))
        file.write("\n")
        file.write("\n")
        processed_data = processed_data[50:]
