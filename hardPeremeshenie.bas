<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Column Modifier</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f4f4f4;
        }

        .container {
            max-width: 600px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            border-radius: 5px;
        }

        h1 {
            text-align: center;
            color: #333;
        }

        input[type="file"] {
            display: block;
            margin: 20px auto;
        }

        button {
            display: block;
            width: 100%;
            padding: 10px;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
        }

        button:disabled {
            background-color: #ccc;
        }

        .result {
            margin-top: 20px;
            text-align: center;
        }

        .result a {
            color: #007bff;
            text-decoration: none;
        }
    </style>
</head>
<body>

<div class="container">
    <h1>Modify Excel Columns (Remove E and G)</h1>
    <input type="file" id="upload" accept=".xlsx" />
    <button id="download" disabled>Download Modified Excel</button>
    <div class="result"></div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
    let workbook;
    
    document.getElementById('upload').addEventListener('change', function (event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                modifyColumns();
                document.getElementById('download').disabled = false;
            };
            reader.readAsArrayBuffer(file);
        }
    });

    function modifyColumns() {
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // Прочитать данные
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Удаление столбцов E (индекс 4) и G (индекс 6 после удаления E)
        for (let i = 0; i < data.length; i++) {
            data[i].splice(6, 1); // Удалить столбец G (индекс 6)
            data[i].splice(4, 1); // Удалить столбец E (индекс 4)
        }

        // Обновляем лист с новыми данными
        const newWorksheet = XLSX.utils.aoa_to_sheet(data);
        workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
    }

    document.getElementById('download').addEventListener('click', function () {
        const newWorkbook = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([newWorkbook], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'modified_excel.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    });
</script>

</body>
</html>
