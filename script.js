document.getElementById('processButton').addEventListener('click', processExcel);

function processExcel() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Please select a file first!');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Get student info from first sheet (Class Master)
        const studentSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Convert to JSON with specific column mapping
        const studentData = XLSX.utils.sheet_to_json(studentSheet, {
            range: 1,  // Skip header row
            header: ['LastName', 'FirstName', 'C', 'Email', 'TeamNumber']
        });
        
        // Clean and validate student data
        const cleanedStudentData = studentData.map(student => ({
            LastName: student.LastName?.trim() || '',
            FirstName: student.FirstName?.trim() || '',
            Email: student.Email?.trim() || '',
            TeamNumber: student.TeamNumber || '',
            ClassParticipation: {} // Will store participation data for each class
        })).filter(student => student.LastName && student.FirstName);
        
        // Process only sheets that start with 'W'
        const classSheets = workbook.SheetNames.slice(1).filter(name => 
            name.toUpperCase().startsWith('W')
        );
        
        classSheets.forEach(sheetName => {
            const classSheet = workbook.Sheets[sheetName];
            processClassSheet(classSheet, sheetName, cleanedStudentData);
        });
        
        // Display preview of combined data
        displayStudentPreview(cleanedStudentData);
        
        // Add summary of processed sheets
        const preview = document.getElementById('preview');
        preview.insertAdjacentHTML('beforeend', 
            `<p>Processed ${classSheets.length} class sheets (Weeks)</p>`);
        
        // Create and download processed file
        createAndDownloadExcel(cleanedStudentData);
    };
    
    reader.readAsArrayBuffer(file);
}

function processClassSheet(sheet, sheetName, studentData) {
    // Convert sheet to JSON, looking for specific columns
    const classData = XLSX.utils.sheet_to_json(sheet, {
        header: ['LastName', 'B', 'C', 'Email', 'E', 'F', 'G', 'Participation'],
        range: sheet.range // Use the sheet's data range
    });

    // Find the row where 'Email' header is located
    const emailHeaderRow = classData.findIndex(row => 
        row.Email && row.Email.toString().toLowerCase().includes('email'));
    
    // Process only rows after the email header
    const relevantData = classData.slice(emailHeaderRow + 1);

    // Initialize all students with 0 participation for this class
    studentData.forEach(student => {
        student.ClassParticipation[sheetName] = '0';
    });

    // Match and update participation data
    relevantData.forEach(classRow => {
        if (!classRow.LastName && !classRow.Email) return;
        
        let matchedStudent = null;

        // First try to match by email (more reliable)
        if (classRow.Email) {
            matchedStudent = studentData.find(s => 
                s.Email.toLowerCase() === classRow.Email.toString().toLowerCase()
            );
        }

        // If no email match, try matching by last name
        if (!matchedStudent && classRow.LastName) {
            matchedStudent = studentData.find(s => 
                s.LastName.toLowerCase() === classRow.LastName.toString().toLowerCase()
            );
        }

        // Update participation if student was found
        if (matchedStudent) {
            matchedStudent.ClassParticipation[sheetName] = classRow.Participation || '0';
        }
    });

    // Log any unmatched students from the class sheet for verification
    console.log(`Unmatched students in ${sheetName}:`);
    relevantData.forEach(classRow => {
        if (!classRow.LastName && !classRow.Email) return;
        
        const matched = studentData.some(s => 
            (classRow.Email && s.Email.toLowerCase() === classRow.Email.toString().toLowerCase()) ||
            (classRow.LastName && s.LastName.toLowerCase() === classRow.LastName.toString().toLowerCase())
        );

        if (!matched) {
            console.log(`- ${classRow.LastName || 'NO_LAST_NAME'} (${classRow.Email || 'NO_EMAIL'})`);
        }
    });
}

function displayStudentPreview(data) {
    const preview = document.getElementById('preview');
    preview.innerHTML = '<h3>Student Information Preview:</h3>';
    
    const table = document.createElement('table');
    table.className = 'student-table';
    
    // Create header row with dynamic class columns
    const headerRow = document.createElement('tr');
    const baseHeaders = ['Last Name', 'First Name', 'Email', 'Team Number'];
    const classHeaders = Object.keys(data[0]?.ClassParticipation || {});
    
    [...baseHeaders, ...classHeaders, 'Participation Average'].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);
    
    // Create data rows (limit to first 5 rows)
    data.slice(0, 5).forEach(student => {
        const tr = document.createElement('tr');
        
        // Add base data
        [student.LastName, student.FirstName, student.Email, student.TeamNumber]
            .forEach(value => {
                const td = document.createElement('td');
                td.textContent = value;
                tr.appendChild(td);
            });
        
        // Add participation data for each class
        classHeaders.forEach(className => {
            const td = document.createElement('td');
            td.textContent = student.ClassParticipation[className] || '';
            tr.appendChild(td);
        });

        // Add average participation
        const participationScores = Object.values(student.ClassParticipation)
            .map(score => parseFloat(score) || 0);
        const averageParticipation = participationScores.length > 0
            ? (participationScores.reduce((a, b) => a + b, 0) / participationScores.length).toFixed(2)
            : '0';
        
        const tdAvg = document.createElement('td');
        tdAvg.textContent = averageParticipation;
        tr.appendChild(tdAvg);
        
        table.appendChild(tr);
    });
    
    preview.appendChild(table);
    preview.insertAdjacentHTML('beforeend', 
        `<p>Total students: ${data.length}</p>`);
}

function createAndDownloadExcel(data) {
    const newWorkbook = XLSX.utils.book_new();
    
    // Prepare data for export
    const exportData = data.map(student => {
        // Calculate average participation
        const participationScores = Object.values(student.ClassParticipation)
            .map(score => parseFloat(score) || 0);
        const averageParticipation = participationScores.length > 0
            ? (participationScores.reduce((a, b) => a + b, 0) / participationScores.length).toFixed(2)
            : '0';

        const row = {
            'Last Name': student.LastName,
            'First Name': student.FirstName,
            'Email': student.Email,
            'Team Number': student.TeamNumber,
            ...student.ClassParticipation,
            'Participation Average': averageParticipation
        };
        return row;
    });
    
    // Convert to worksheet
    const newWorksheet = XLSX.utils.json_to_sheet(exportData);
    
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Class Master");
    
    // Generate Excel file
    XLSX.writeFile(newWorkbook, "processed_student_data.xlsx");
} 