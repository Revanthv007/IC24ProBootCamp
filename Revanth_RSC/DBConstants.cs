using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revanth_RSC
{
    internal class DBConstants
    {
        public static string INSTERPRODUCT =  @"IF NOT EXISTS(SELECT * FROM ProductsMaster WHERE ProductCode = @ProductCode)
                                BEGIN
	                                INSERT ProductsMaster (ProductCode, ProductName, PricePerUnit)
	                                SELECT @ProductCode,@ProductName, @PricePerUnit
                                END";

        public static string INSERTSTOCK = @"
                IF EXISTS (SELECT * FROM Stock WHERE [Date] = CAST(@Date AS DATE) AND ProductIdFk =(SELECT ProductIdPk FROM ProductsMaster WHERE ProductCode = @ProductCode ))
                BEGIN
                    UPDATE Stock 
                    SET QuantityAvailable = @Quantity 
                    WHERE [Date] = CAST(@Date AS DATE) AND ProductIdFk =(SELECT ProductIdPk FROM ProductsMaster WHERE ProductCode = @ProductCode )
                END
                ELSE
                BEGIN
                    INSERT INTO Stock (ProductIdFk, StoreIdFk, [Date], QuantityAvailable,PricePerUnit)
                    VALUES ((SELECT ProductIdPk FROM ProductsMaster WHERE ProductCode = @ProductCode ), (Select StoreIdPk FROM Stores WHERE StoreCode = @StoreCode), CAST(@Date AS DATE), @Quantity,@PricePerUnit)
                END
            ";

        public static string INSERTORUPDATEEMP =@"
                                IF EXISTS (SELECT * FROM Employees WHERE EmpCode = @EmployeeCode)
                                BEGIN
                                    UPDATE Employees 
                                    SET 
                                        EmployeeName = @EmployeeName,
                                        [Role] = @Role,
                                        DateOfJoining = CAST(@DOJ AS DATE),
                                        DateOfLeaving = CAST(@DOL AS DATE),
                                        ContactNumber = @ContactNumber,
                                        Gender = @Gender,
                                        Salary = @Salary,
                                        StoreIdFk = (SELECT StoreIdPk FROM Stores WHERE StoreCode = @StoreCode)
                                    WHERE EmpCode = @EmployeeCode
                                END
                                ELSE
                                BEGIN
                                    INSERT INTO Employees 
                                        (EmpCode, EmployeeName, Role, DateOfJoining, DateOfLeaving, ContactNumber, Gender, Salary, StoreIdFk)
                                    VALUES 
                                        (@EmployeeCode, @EmployeeName, @Role, CAST(@DOJ AS DATE), CAST(@DOL AS DATE), @ContactNumber, @Gender, @Salary, (SELECT StoreIdPk FROM Stores WHERE StoreCode = @StoreCode))
                                END";

        public static string STOREUPDATE = $"Update Stores Set StoreName = @StoreName, Location = @Location, ManagerName= @Manager, ContactNumber = @ContactNumber where StoreCode = @StoreCode";
    }

    public enum FileProcessStatus
    {
        Processed,
        Archive
    }
}
