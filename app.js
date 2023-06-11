window.addEventListener('DOMContentLoaded', function() {
  // Initialize Flatpickr for the month picker field
  flatpickr('#monthPicker', {
    dateFormat: 'F Y',
    locale: 'ru',
  });

  // Initialize Flatpickr for the ad days picker field
  flatpickr('#adDaysPicker', {
    mode: 'multiple',
    dateFormat: 'd',
    locale: 'ru',
  });

  // Generate Table
  document.getElementById('generateTableBtn').addEventListener('click', function() {
    const month = document.getElementById('monthPicker').value;
    const customer = document.getElementById('customerInput').value;
    const adDays = document.getElementById('adDaysPicker').value.split(',');
    const adTimes = Array.from(document.querySelectorAll('#adTimesSelect option:checked')).map(option => option.value);
    const duration = document.getElementById('durationInput').value;
    const video = document.getElementById('videoTitle').value;

    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();

    // Fetch the template file
    fetch('template.xlsx')
      .then(function(response) {
        return response.arrayBuffer();
      })
      .then(function(data) {
        // Load the template into the workbook
        return workbook.xlsx.load(data);
      })
      .then(function() {
        // Modify the workbook as needed
        const worksheet = workbook.getWorksheet(1);

        // Set the values
        worksheet.getCell('F4').value = customer;
        worksheet.getCell('F5').value = video;
        worksheet.getCell('F6').value = duration;
        worksheet.getCell('I1').value = month;

        // Set the ad days and ad times values
        adDays.forEach(function(day) {
          const column = String.fromCharCode(65 + parseInt(day)); // Convert day to column letter
          adTimes.forEach(function(time, index) {
            let row;
            if (time === '08:10') {
              row = 11;
            } else if (time === '08:40') {
              row = 12;
            } else if (time === '09:10') {
              row = 13;
            } else if (time === '14:30') {
              row = 14;
            } else if (time === '17:10') {
              row = 15;
            } else if (time === '21:05') {
              row = 16;
            }
            const cell = worksheet.getCell(column + row);
            cell.value = duration;
          });
        });         

        // Generate the formatted Excel file
        return workbook.xlsx.writeBuffer();
      })
      .then(function(buffer) {
        // Create a Blob from the buffer
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Generate a temporary download link
        const url = window.URL.createObjectURL(blob);

        // Create a download link element
        const link = document.createElement('a');
        link.href = url;

        const currentDate = new Date();
        const formattedDate = currentDate.toISOString().split('T')[0];
        const filename = `Медиаплан-${video}-${formattedDate}.xlsx`;
        
        link.download = filename;

        // Trigger the download
        link.click();

        // Clean up the temporary download link
        window.URL.revokeObjectURL(url);
      })
      .catch(function(error) {
        console.error('Error:', error);
      });
  });
});