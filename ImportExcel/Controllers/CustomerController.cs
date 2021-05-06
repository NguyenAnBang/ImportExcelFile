using Dapper;
using ImportExcel.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MySqlConnector;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ImportExcel.Controllers
{
    [Route("api/[controller]s")]
    [ApiController]
    public class CustomerController : ControllerBase
    {
        IDbConnection dbConnection;

        string connectionString = ""
                + "Host=47.241.69.179;"
                + "Port=3306;"
                + "User Id=dev;"
                + "Password=12345678;"
                + "Database = MF825_Import_NABANG;"
                + "convert zero datetime=True";
        string sqlCommand = "";
        DynamicParameters dynamicParameters = new DynamicParameters();

        [HttpPost("Import")]      
        public async Task<IActionResult> Import(IFormFile formFile, CancellationToken cancellationToken)
        {
            
            if (formFile == null || formFile.Length <= 0)
            {
                //Tạo thông báo lỗi
                var errorMsg = new
                {
                    devMsg = Properties.Resources.devMsg_NoFile,
                    userMsg = Properties.Resources.userMsg
                };
                return BadRequest(errorMsg);
            }
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                //Tạo thông báo lỗi
                var errorMsg = new
                {
                    devMsg = Properties.Resources.devMsg_InvalidFile,
                    userMsg = Properties.Resources.userMsg
                };
                return BadRequest(errorMsg);
            }
            //Tạo 1 mảng để nhận giá trị
            var customerList = new List<Customer>();
            
            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 3; row <= rowCount; row++)
                    {
                        //Vì dateOfBirth là DateTime nên không thể gán trực tiếp với string
                        //Gán giá trị của ngày sinh trong file excel cho 1 biến string
                        //Format biến string này trả về datetime
                        //Gán giá trị datetime trả về cho dateOfBirth
                        string dateOfBirthFormat = worksheet.Cells[row, 6].Value == null ? "" : worksheet.Cells[row, 6].Value.ToString().Trim();
                        customerList.Add(new Customer
                        {
                            customerCode = worksheet.Cells[row, 1].Value == null ? "" : worksheet.Cells[row, 1].Value.ToString().Trim(),
                            customerFullName = worksheet.Cells[row, 2].Value == null ? "" : worksheet.Cells[row, 2].Value.ToString().Trim(),
                            memberCardCode = worksheet.Cells[row, 3].Value == null ? "" : worksheet.Cells[row, 3].Value.ToString().Trim(),
                            customerGroupName = worksheet.Cells[row, 4].Value == null ? "" : worksheet.Cells[row, 4].Value.ToString().Trim(),
                            phoneNumber = worksheet.Cells[row, 5].Value == null ? "" : worksheet.Cells[row, 5].Value.ToString().Trim(),
                            dateOfBirth = dobFormat(dateOfBirthFormat),
                            companyName = worksheet.Cells[row, 7].Value == null ? "" : worksheet.Cells[row, 7].Value.ToString().Trim(),
                            taxCode = worksheet.Cells[row, 8].Value == null ? "" : worksheet.Cells[row, 8].Value.ToString().Trim(),
                            email = worksheet.Cells[row, 9].Value == null ? "" : worksheet.Cells[row, 9].Value.ToString().Trim(),
                            address = worksheet.Cells[row, 10].Value == null ? "" : worksheet.Cells[row, 10].Value.ToString().Trim(),
                            note = worksheet.Cells[row, 11].Value == null ? "" : worksheet.Cells[row, 11].Value.ToString().Trim(),
                            customerStatus = ""
                        }); 
                    }
                }
            }

            //Validate dữ liệu

            //Check trùng trong tệp nhập khẩu
            //Check trùng mã khách hàng trong tệp
            for (int i = 0; i < customerList.Count; i++)
            {
                for (int j = i+1; j < customerList.Count; j++)
                {
                    if (customerList[i].customerCode == customerList[j].customerCode)
                    {
                        customerList[j].customerStatus += Properties.Resources.File_CustomerCode_Exists;
                        break;
                    }                  
                }                
            }
            //Check trùng số điện thoại trong tệp
            for (int i = 0; i < customerList.Count; i++)
            {
                for (int j = i+1; j < customerList.Count; j++)
                {
                    if (customerList[i].phoneNumber == customerList[j].phoneNumber)
                    {
                        customerList[j].customerStatus += Properties.Resources.File_PhoneNumber_Exists;
                        break;
                    }
                }                
            }
            //Check trùng email trong tệp
            for (int i = 0; i < customerList.Count; i++)
            {
                for (int j = i+1; j < customerList.Count; j++)
                {
                    if (customerList[i].email == customerList[j].email)
                    {
                        customerList[j].customerStatus += Properties.Resources.File_Email_Exists;
                        break;
                    }
                }                
            }
            //Check trùng trong database
            foreach (var customer in customerList)
            {
                //Check trùng mã khách hàng trong database                   
                if (CheckExists("CustomerCode", customer.customerCode)) customer.customerStatus += Properties.Resources.Database_CustomerCode_Exists;
                //Check trùng số điện thoại trong database
                if (CheckExists("PhoneNumber", customer.phoneNumber)) customer.customerStatus += Properties.Resources.Database_PhoneNumber_Exists;
                //Check trùng email trong database
                if (CheckExists("Email", customer.email)) customer.customerStatus += Properties.Resources.Database_Email_Exists;
                //Check nhóm khách hàng có trong hệ thống hay không
                CheckCustomerGroupNameExists(customer);
            }
            //Kiểm tra dữ liệu có hợp lệ không
            foreach (var customer in customerList)
            {
                if (customer.customerStatus == "") customer.customerStatus = Properties.Resources.Accept;
            }


            //Push dữ liệu lên database
            sqlCommand = "Proc_InsertCustomer";
            //Kiểm tra số dòng bị ảnh hưởng
            var rowsAffect = 0;
            //Thông báo
            var Noti = "";
            //Tạo mảng để nhận giá trị đẩy lên database thành công
            var successList = new List<Customer>();
            for (int i = 0; i < customerList.Count; i++)
            {
                if(customerList[i].customerStatus == "Hợp lệ")
                {
                    using (dbConnection = new MySqlConnection(connectionString))
                    {
                        rowsAffect += dbConnection.Execute(sqlCommand, param: customerList[i], commandType: CommandType.StoredProcedure);
                    }
                    
                    successList.Add(customerList[i]);
                }
                
            }
            if(successList.Count > 0) Noti = Properties.Resources.Success;
            else Noti = Properties.Resources.Failure;
            var response = new
            {
                Status = Noti,
                numberOfSuccessRow = rowsAffect,
                numberOfFailRow = customerList.Count - rowsAffect,
                //data = successList //Hiện ra danh sách dữ liệu được push thành công
                data = customerList //Hiện ra danh sách tất cả dữ liệu
            };

            if (rowsAffect > 0)
            {
                return Ok(response);
            }
            else
            {
                return BadRequest(response);
            }

        }
        /// <summary>
        /// Format ngày sinh từ dạng string sang datetime
        /// </summary>
        /// <param name="dob"></param>
        /// <returns></returns>
        DateTime? dobFormat(string dob)
        {
            if (dob == "") return null;
            //tách ký tự '/', các phần tử dd,mm,yyyy được nhét vào một mảng string
            string[] dobItems = dob.Split('/');
            //Xét mảng dobItems có 1 phần tử, tức là chỉ có yyyy
            if (dobItems.Length == 1) dob = "01/01/" + dob;
            //Xét mảng dobItems có 2 phần tử, tức là có MM/yyyy
            else if (dobItems.Length == 2) dob = "01/" + dob;
            
            //Return giá trị Datetime
            DateTime outputDateTimeValue;
            outputDateTimeValue = DateTime.ParseExact(dob, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);
            return outputDateTimeValue;

        }
        /// <summary>
        /// Check trùng trong database
        /// </summary>
        /// <param name="propertyName"></param>
        /// <param name="value"></param>
        /// <returns>True: Nếu trùng</returns>
        /// <returns>False: Nếu không trùng</returns>
        private bool CheckExists(string propertyName, string value)
        {
            using (dbConnection = new MySqlConnection(connectionString))
            {
                sqlCommand = $"Proc_Check{propertyName}Exists";
                dynamicParameters.Add($"{propertyName}", value);
                var exists = dbConnection.QueryFirstOrDefault<bool>(sqlCommand, param: dynamicParameters, commandType: CommandType.StoredProcedure);
                return exists;
            }          
        }
        /// <summary>
        /// Kiểm tra xem tên nhóm khách hàng có tồn tại hay không, nếu có thì gán customerGroupID = id của nhóm khách hàng tương ứng
        /// </summary>
        /// <param name="customer"></param>
        private void CheckCustomerGroupNameExists(Customer customer)
        {
            using (dbConnection = new MySqlConnection(connectionString))
            {
                sqlCommand = "Proc_GetCustomerGroupByName";
                dynamicParameters.Add("customerGroupName", customer.customerGroupName);
                var cgName = dbConnection.QueryFirstOrDefault<CustomerGroup>(sqlCommand, param: dynamicParameters, commandType: CommandType.StoredProcedure);
                if (cgName == null)
                {
                    customer.customerStatus += Properties.Resources.Database_CustomerGroup_NotExists;
                }
                else
                {
                    customer.customerGroupID = cgName.CustomerGroupId;
                }
            }
        }




    }   
}
