namespace Terrasoft.Configuration
{
	using System;
	using System.IO;
	using System.Linq;
	using System.Drawing;
	using System.Collections.Generic;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Entities.Events;
	using OfficeOpenXml;
	using OfficeOpenXml.Style;
	using OfficeOpenXml.Table;
	using OfficeOpenXml.FormulaParsing.ExcelUtilities;

	public class AprExcelHelper
	{
		private Guid White = new Guid("0a8d701b-7c35-4948-b5c9-32e035f52eb3");
		private Guid Green = new Guid("9ea18075-76b1-4585-bd13-4e6eea1824a3");
		private Guid Yellow = new Guid("da5be7ab-fe4f-4ea8-a982-4de3423022b1");
		private Guid Red = new Guid("39201f8f-772f-4556-b406-1775e9929999");
		private Guid Incoming = new Guid("7f9d1f86-f36b-1410-068c-20cf30b39373");


		private UserConnection UserConnection { get; set; }

		public AprExcelHelper(UserConnection userConnection)
		{
			UserConnection = userConnection;
		}

		private ExcelPackage ConverByteArrayTOExcepPackage(byte[] bytes)
		{
			using (MemoryStream memStream = new MemoryStream(bytes))
			{
				ExcelPackage package = new ExcelPackage(memStream);
				return package;
			}
		}

		public ExcelPackage GetTemplatePackage()
		{
			var template = (byte[])Terrasoft.Core.Configuration.SysSettings.GetValue(UserConnection, "AprAuthoVerificationTemplate");
			var package = ConverByteArrayTOExcepPackage(template);
			return package;
		}


		public void SaveData(byte[] bytes)
		{
			var schema = UserConnection.EntitySchemaManager.GetInstanceByName("ContactFile");
			var entity = schema.CreateEntity(UserConnection);
			Guid callFileId = Guid.NewGuid();
			entity.SetDefColumnValues();
			entity.SetColumnValue("Id", callFileId);
			entity.SetColumnValue("Name", "test.xlsx");
			entity.SetColumnValue("TypeId", new Guid("529bc2f8-0ee0-df11-971b-001d60e938c6"));
			entity.SetColumnValue("ContactId", new Guid("10f6632c-05d0-9575-15a7-4e8ef0d42915"));
			entity.SetColumnValue("Data", bytes);
			entity.Save();
		}

		public void Create()
		{
			var package = GetTemplatePackage();
			var entities = GetData();
			SetExcelData(package, entities);
			var bytes = package.GetAsByteArray();
			SaveData(bytes);
		}

		public EntityCollection GetData()
		{
			var esq = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "ArpVerificationDeliveries");
			esq.AddAllSchemaColumns();
			var entities = esq.GetEntityCollection(UserConnection);
			return entities;
		}

		private void SetRangeBorders(ExcelWorksheet sheet, int rowStart, int columnStart, int rowEnd, int columnEnd)
		{
			using (ExcelRange border = sheet.Cells[rowStart, columnStart, rowEnd, columnEnd])
			{
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Top.Color.SetColor(Color.LightBlue);
				border.Style.Border.Bottom.Color.SetColor(Color.LightBlue);
				border.Style.Border.Left.Color.SetColor(Color.LightBlue);
				border.Style.Border.Right.Color.SetColor(Color.LightBlue);
			}
		}

		public void SetExcelData(ExcelPackage package, EntityCollection entities)
		{
			var firstRow = 10;
			var fistCol = 1;
			var lastCol = 16;
			var sheet = package.Workbook.Worksheets[2];
			var position = 1;

			SetRangeBorders(sheet, firstRow - 1, fistCol, firstRow + entities.Count - 1, lastCol);

			var i = 1;
			var row = firstRow;
			foreach (var entity in entities)
			{
				var col = 1;
				if (i < entities.Count)
				{
					sheet.Cells[row, fistCol, row, lastCol].Copy(sheet.Cells[row + 1, fistCol, row + 1, lastCol]);
				}
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<Guid>("Id");
				sheet.Cells[row, col++].Value = position++;
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<string>("AprDebtorName");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<string>("AprCustomerName");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<DateTime>("AprShippingDate");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<DateTime>("AprDocumentDate");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<string>("AprNumber");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<string>("AprInvoiceNumber");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<double>("AprAmount");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<double>("AprPaidUp");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<double>("AprDebtAmount");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<int>("AprPostponingDays");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<int>("AprDelay");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<DateTime>("AprDelayUntil");
				sheet.Cells[row, col++].Value = entity.GetTypedColumnValue<string>("AprComment");
				row++;
				i++;
			}
		}

		public bool GetIsIncoming(Guid activityId)
		{
			var activity = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "Activity");
			activity.AddColumn("MessageType");
			var entity = activity.GetEntity(UserConnection, activityId);
			var messageType = entity.GetTypedColumnValue<Guid>("MessageTypeId");
			return (messageType == Incoming);
		}

		public void Read(Guid activityId, Guid activityFileId)
		{
			if (!GetIsIncoming(activityId))
			{
				return;
			}
			var bytes = GetActivityFile(activityFileId);
			var package = ConverByteArrayTOExcepPackage(bytes);
			var result = ReadColumnExcel(package);
			SetActivityResult(activityId, result);
		}

		public byte[] GetActivityFile(Guid activityFileId)
		{
			EntitySchema schema = UserConnection.EntitySchemaManager.GetInstanceByName("ActivityFile");
			EntitySchemaQuery esq = new EntitySchemaQuery(schema);
			esq.AddAllSchemaColumns();
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, "Id", activityFileId));

			EntityCollection activityFileEntities = esq.GetEntityCollection(UserConnection);
			var data = new byte[0];
			foreach (Entity activityFile in activityFileEntities)
			{
				var name = activityFile.GetTypedColumnValue<string>("Name");
				data = activityFile.GetBytesValue("Data");
			}
			return data;
		}

		public AprExcelResult ReadColumnExcel(ExcelPackage package)
		{
			try
			{
				var sheet = package.Workbook.Worksheets[2];
				var firstRow = 10;
				var lastRow = sheet.Dimension.End.Row;

				var result = new AprExcelResult()
				{
					ConfirmedIds = new List<Guid>(),
					NegativeIds = new List<Guid>(),
					TotalCount = 0,
					EmptyCount = 0
				};
				List<string> confirmedComments = new List<string>();
				List<string> negativeComments = new List<string>();

				for (var row = firstRow; row <= lastRow; row++)
				{
					if (sheet.Cells[row, 15].Value.ToString() == "Подтверждаю")
					{
						result.ConfirmedIds.Add(new Guid(sheet.Cells[row, 1].Text));
						if (!String.IsNullOrEmpty(sheet.Cells[row, 16].Text))
						{
							confirmedComments.Add(sheet.Cells[row, 7].Text + " - " + sheet.Cells[row, 16].Text);
						}
					}
					else
					if (sheet.Cells[row, 15].Value.ToString() == "Не подтверждаю")
					{
						result.NegativeIds.Add(new Guid(sheet.Cells[row, 1].Text));
						if (!String.IsNullOrEmpty(sheet.Cells[row, 16].Text))
						{
							negativeComments.Add(sheet.Cells[row, 7].Text + " - " + sheet.Cells[row, 16].Text);
						}
					}
					else
					{
						result.EmptyCount++;
					}
					result.TotalCount++;
				}
				result.ConfirmedCount = result.ConfirmedIds.Count;
				result.NegativeCount = result.NegativeIds.Count;
				result.ConfirmedComments = String.Join("\r\n", confirmedComments.ToArray());
				result.NegativeComments = String.Join("\r\n", negativeComments.ToArray());
				return result;
			}
			catch (Exception e)
			{
				return null;
			}
		}

		public void SetActivityResult(Guid activityId, AprExcelResult result)
		{

			var activity = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "Activity");
			activity.AddAllSchemaColumns();
			var entity = activity.GetEntity(UserConnection, activityId);
			var colorStatusId = GetActivityColorStatus(result);
			if (entity == null)
			{
				return;
			}
			entity.SetColumnValue("AprIndicatorId", colorStatusId);
			entity.SetColumnValue("DetailedResult", result.ConfirmedComments + "\r\n" + result.NegativeComments);
			entity.Save();
		}

		public Guid GetActivityColorStatus(AprExcelResult result)
		{
			var colorStatusId = Guid.Empty;
			if (result.TotalCount == result.ConfirmedCount)
			{
				colorStatusId = Green;
			}
			else
			if (result.ConfirmedCount >= 0)
			{
				colorStatusId = Yellow;
			}
			else
			{
				colorStatusId = Red;
			}
			return colorStatusId;
		}
	}

	public class AprExcelResult
	{
		public int TotalCount;
		public int ConfirmedCount;
		public int NegativeCount;
		public int EmptyCount;

		public string ConfirmedComments;
		public string NegativeComments;

		public List<Guid> ConfirmedIds;
		public List<Guid> NegativeIds;
	}
}