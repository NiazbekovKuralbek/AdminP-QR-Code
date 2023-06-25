using Microsoft.AspNetCore.Mvc;
using AdminP_QR_Code_.Models;
using Newtonsoft.Json;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading.Tasks;

namespace AdminP_QR_Code_.Controllers
{
    //Здесь был я

    [ApiController]
    [Route("api/pointsale")]
    public class PointSaleController : ControllerBase
    {
       
        [HttpPost]
        public IActionResult CreatePointSale([FromBody] PointSale pointSale)
        {
            // Выполнение операций с полученными данными pointSale
            

            // Преобразование ссылки в QR-код
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(pointSale.qr_data, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20); // 20 - размер точки QR-кода

            
            string tempFilePath = Path.GetTempFileName();
            qrCodeImage.Save(tempFilePath, ImageFormat.Png);// Сохранение QR-кода во временный файл .PNG

            
            string existingFilePath = "C:\\Ohagi\\Studyy\\3rd year\\6rd semester\\Практика\\AdminP(QR Code)\\QR_Code.docx";
            Application wordApp = new Application();
            Document wordDoc = wordApp.Documents.Open(existingFilePath);

            // Вставка QR-кода в документ Word
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Range range = wordDoc.Range();
            range.InlineShapes.AddPicture(tempFilePath, ref missing, ref missing, ref missing);

            // Сохранение документа Word
            wordDoc.Save();
            wordDoc.Close();
            wordApp.Quit();

            // Удаление временного файла
            System.IO.File.Delete(tempFilePath);

            // Возвращение файла в качестве ответа
            var fileStream = new FileStream(existingFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return new FileStreamResult(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = "modified_document.docx"
            };
        }

        //[HttpPost]
        //public IActionResult Create(TaskModel model)
        //{
        //    // Проверка на валидность модели данных
        //    if (!ModelState.IsValid)
        //    {
        //        return BadRequest(ModelState);
        //    }


        //    // Создание новой записи
        //    var newTask = new TaskModel
        //    {
        //        Id = Guid.NewGuid().ToString(), // Генерация уникального идентификатора
        //        Title = model.Title,
        //        Description = model.Description,
        //        // Другие свойства модели данных
        //    };

        //    // Добавление новой записи в коллекцию
        //    tasks.Add(newTask);

        //    // Возврат результата операции с кодом 201 Created
        //    return CreatedAtAction(nameof(GetById), new { id = newTask.Id }, newTask);
        //}

        //[HttpDelete("{id}")]
        //public IActionResult Delete(string id)
        //{
        //    // Поиск записи по идентификатору
        //    var task = tasks.FirstOrDefault(t => t.Id == id);
        //    if (task == null)
        //    {
        //        return NotFound(); // Если запись не найдена, возвращаем статус 404 Not Found
        //    }

        //    // Удаление записи из коллекции
        //    tasks.Remove(task);

        //    // Возврат результата операции с кодом 204 No Content
        //    return NoContent();
        //}


        //[HttpPut("{id}")]

        //public IActionResult Update(string id, TaskModel model)
        //{
        //    // Проверка на валидность модели данных
        //    if (!ModelState.IsValid)
        //    {
        //        return BadRequest(ModelState);
        //    }

        //    // Поиск записи по идентификатору
        //    var task = tasks.FirstOrDefault(t => t.Id == id);
        //    if (task == null)
        //    {
        //        return NotFound(); // Если запись не найдена, возвращаем статус 404 Not Found
        //    }

        //    // Обновление свойств записи
        //    task.Title = model.Title;
        //    task.Description = model.Description;
        //    // Обновление других свойств модели данных

        //    // Возврат результата операции с кодом 200 OK
        //    return Ok(task);
        //}
        //[HttpGet("{id}")]
        //public IActionResult GetById(string id)
        //{
        //    var task = tasks.FirstOrDefault(t => t.Id == id);
        //    if (task == null)
        //    {
        //        return NotFound();
        //    }

        //    return Ok(task);
        //}

        //// Другие методы контроллера (например, получение списка всех задач, получение задачи по идентификатору и т. д.)
    }
}
