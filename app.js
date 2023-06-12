window.addEventListener('DOMContentLoaded', function() {
  // Инициализация Flatpickr для поля с выбором месяца
  flatpickr('#monthPicker', {
    dateFormat: 'F Y',
    locale: 'ru',
  });

  // Инициализация Flatpickr для поля с выбором дней
  flatpickr('#adDaysPicker', {
    mode: 'multiple',
    dateFormat: 'd',
    locale: 'ru',
  });

  // Генерация таблицы
  document.getElementById('generateTableBtn').addEventListener('click', function() {
    const month = document.getElementById('monthPicker').value;
    const customer = document.getElementById('customerInput').value;
    const adDays = document.getElementById('adDaysPicker').value.split(',');
    const adTimes = Array.from(document.querySelectorAll('#adTimesSelect option:checked')).map(option => option.value);
    const duration = document.getElementById('durationInput').value;
    const video = document.getElementById('videoTitle').value;

    // Создание Excel-книги
    const workbook = new ExcelJS.Workbook();

    // Получение файла шаблона
    fetch('template.xlsx')
      .then(function(response) {
        return response.arrayBuffer();
      })
      .then(function(data) {
        // Загрузка шаблона в книгу
        return workbook.xlsx.load(data);
      })
      .then(function() {
        // Изменение книги по мере необходимости
        const worksheet = workbook.getWorksheet(1);

        // Выбор нужных ячеек
        worksheet.getCell('F4').value = customer;
        worksheet.getCell('F5').value = video;
        worksheet.getCell('F6').value = duration;
        worksheet.getCell('I1').value = month;

        // Установка значений от времени и дней
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

        // Генерация оформленного Excel-файла
        return workbook.xlsx.writeBuffer();
      })
      .then(function(buffer) {
        // Create a Blob from the buffer
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Генерация временной ссылки на скачивание
        const url = window.URL.createObjectURL(blob);

        // Создание элемента для ссылки на скачивание
        const link = document.createElement('a');
        link.href = url;

        const currentDate = new Date();
        const formattedDate = currentDate.toISOString().split('T')[0];
        const filename = `Медиаплан-${video}-${formattedDate}.xlsx`;
        
        link.download = filename;

        // Триггер загрузки
        link.click();

        // Очистка временной ссылки на скачивание
        window.URL.revokeObjectURL(url);
      })
      .catch(function(error) {
        console.error('Ошибка: ', error);
      });
  });
});