namespace Terrasoft.Configuration
{
	using System;
	using System.Linq;
	using System.Collections.Generic;
	using System.Runtime.Serialization;
	using System.Text.RegularExpressions;

	using Newtonsoft.Json;

	using SugarRestSharp;
	using Sugar = SugarRestSharp.Models;

	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Entities.Events;

	public class AprSugarAccountIntegration
	{
		private string UncorrectId;
		private string SugarURL;
		private string AdminSugarLogin;
		private string AdminSugarPassword;
		private string AdminSugarUserId;
		private Guid CommunicationTypeId = new Guid("6a3fb10c-67cc-df11-9b2a-001d60e938c6");
		private Guid EmailTypeId = new Guid("ee1c85c3-cfcb-df11-9b2a-001d60e938c6");
		private Entity AccountEntity;
		private EntityCollection AccountEntityCollection;
		private UserConnection UserConnection { get; set; }

		public AprSugarAccountIntegration(UserConnection userConnection)
		{
			UserConnection = userConnection;
			PrepareSysSettings();
		}

		private void PrepareSysSettings()
		{
			AdminSugarLogin = (string)Terrasoft.Core.Configuration.SysSettings.GetValue(UserConnection, "AprSugarAdminLogin");
			AdminSugarPassword = (string)Terrasoft.Core.Configuration.SysSettings.GetValue(UserConnection, "AprSugarAdminPassword");
			AdminSugarUserId = (string)Terrasoft.Core.Configuration.SysSettings.GetValue(UserConnection, "AprSugarAdminUserId");
			SugarURL = (string)Terrasoft.Core.Configuration.SysSettings.GetValue(UserConnection, "AprSugarURL");
		}

		private void GetAccountEntity(Guid AccountId)
		{
			var esq = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "Account");
			esq.AddAllSchemaColumns();
			// Отключение механизма выборки локализуемых данных.
			esq.UseLocalization = false;
			esq.AddColumn("Type.Name");
			esq.AddColumn("EmployeesNumber.Name");
			esq.AddColumn("Industry.Name");
			AccountEntity = esq.GetEntity(UserConnection, AccountId);
		}

		private string GetSugarAccountId(SugarRestResponse response)
		{
			if (response.Data == null)
			{
				return string.Empty;
			}
			List<Sugar.Account> accounts = (List<Sugar.Account>)response.Data;
			if (accounts.Count == 0)
			{
				return string.Empty;
			}
			return accounts[0].Id;
		}

		private string GetSugarOwnertId()
		{
			var ownerId = AccountEntity.GetTypedColumnValue<Guid>("OwnerId");
			var esq = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "AprDataAuthorization");
			esq.AddAllSchemaColumns();
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, "AprContactId", ownerId));
			var entity = esq.GetEntityCollection(UserConnection).FirstOrDefault();
			if (entity == null)
			{
				return AdminSugarUserId;
			}
			return entity.GetTypedColumnValue<string>("AprSugarOwnerId");
		}

		private void UpdateSugarAccountId(Guid AccountId, string sugarId)
		{
			var update = new Update(UserConnection, "Account")
			.Set("AprSugarId", Column.Parameter(sugarId))
			.Where("Id").IsEqual(Column.Parameter(AccountId));
			update.Execute();
		}

		public void CreateAccount(Guid AccountId)
		{
			GetAccountEntity(AccountId);
			var sugarCrmId = ReadEntryListResponse();
			if (sugarCrmId != string.Empty)
			{
				UpdateModule(sugarCrmId);
			}
			else
			{
				sugarCrmId = CreateModule();
			}
			UpdateSugarAccountId(AccountId, sugarCrmId);
		}

		public void UpdateAccount(Guid AccountId)
		{
			GetAccountEntity(AccountId);
			var sugarCrmId = AccountEntity.GetTypedColumnValue<string>("AprSugarId");
			UpdateModule(sugarCrmId);
		}

		private void GetPhoneOfficeAmount()
		{
			var esq = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "AccountCommunication");
			var number = esq.AddColumn("Number");
			esq.AddColumn("CommunicationType");
			number.OrderByDesc();
			var esqAccountId = esq.CreateFilterWithParameters(FilterComparisonType.Equal, "Account", AccountEntity.GetTypedColumnValue<Guid>("Id"));
			var esqTypeId = esq.CreateFilterWithParameters(FilterComparisonType.Equal, "CommunicationType", CommunicationTypeId);
			esq.Filters.Add(esqAccountId);
			esq.Filters.Add(esqTypeId);
			AccountEntityCollection = esq.GetEntityCollection(UserConnection);
		}

		private string GetAccountEmail()
		{
			var esq = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "AccountCommunication");
			var number = esq.AddColumn("Number");
			esq.AddColumn("CommunicationType");
			number.OrderByDesc();
			var esqAccountId = esq.CreateFilterWithParameters(FilterComparisonType.Equal, "Account", AccountEntity.GetTypedColumnValue<Guid>("Id"));
			var esqTypeId = esq.CreateFilterWithParameters(FilterComparisonType.Equal, "CommunicationType", EmailTypeId);
			esq.Filters.Add(esqAccountId);
			esq.Filters.Add(esqTypeId);
			AccountEntityCollection = esq.GetEntityCollection(UserConnection);
			if (AccountEntityCollection.Count == 0)
			{
				return string.Empty;
			}
			return AccountEntityCollection[0].GetTypedColumnValue<string>("Number");
		}

		private string CreateModule()
		{
			var ownerId = GetSugarOwnertId();

			string moduleName = "Accounts";

			var client = new SugarRestClient(SugarURL, AdminSugarLogin, AdminSugarPassword);
			var request = new SugarRestRequest(moduleName, RequestType.Create);

			var AccountToCreate = new Sugar.Account();

			if (String.IsNullOrEmpty(ownerId) || ownerId == Guid.Empty.ToString())
			{
				AccountToCreate.AssignedUserId = AdminSugarUserId;
			}
			else
			{
				AccountToCreate.AssignedUserId = ownerId;
			}

			GetPhoneOfficeAmount();
			AccountToCreate.PhoneFax = "";
			if (AccountEntityCollection.Count == 0)
			{
				AccountToCreate.PhoneOffice = "";
			}
			else if (AccountEntityCollection.Count >= 2)
			{
				AccountToCreate.PhoneOffice = AccountEntityCollection[0].GetTypedColumnValue<string>("Number");
				AccountToCreate.PhoneFax = AccountEntityCollection[1].GetTypedColumnValue<string>("Number");
			}
			else
			{
				AccountToCreate.PhoneOffice = AccountEntityCollection[0].GetTypedColumnValue<string>("Number");
				AccountToCreate.PhoneFax = "";
			}

			AccountToCreate.Name = AccountEntity.GetTypedColumnValue<string>("Name");
			AccountToCreate.AccountType = AccountEntity.GetTypedColumnValue<string>("Type_Name");

			AccountToCreate.Industry = AccountEntity.GetTypedColumnValue<string>("Industry_Name");
			AccountToCreate.Employees = AccountEntity.GetTypedColumnValue<string>("EmployeesNumber_Name");
			AccountToCreate.Website = AccountEntity.GetTypedColumnValue<string>("Web");

			var email = GetAccountEmail();
			AccountToCreate.MainEmailAddress = email;

			request.Parameter = AccountToCreate;

			List<string> selectFields = new List<string>();
			selectFields.Add(nameof(Sugar.Account.AssignedUserId));
			selectFields.Add(nameof(Sugar.Account.Name));
			selectFields.Add(nameof(Sugar.Account.AccountType));
			selectFields.Add(nameof(Sugar.Account.PhoneOffice));
			selectFields.Add(nameof(Sugar.Account.PhoneFax));
			selectFields.Add(nameof(Sugar.Account.Industry));
			selectFields.Add(nameof(Sugar.Account.Employees));
			selectFields.Add(nameof(Sugar.Account.Website));
			selectFields.Add(nameof(Sugar.Account.MainEmailAddress));

			request.Options.SelectFields = selectFields;

			SugarRestResponse response = client.Execute(request);
			return (string)response.Data;
		}

		private void UpdateModule(string sugarCrmId)
		{
			var ownerId = GetSugarOwnertId();
			var client = new SugarRestClient(SugarURL, AdminSugarLogin, AdminSugarPassword);
			var readRequest = new SugarRestRequest("Accounts", RequestType.ReadById);
			string AccountId = sugarCrmId;
			readRequest.Parameter = AccountId;
			SugarRestResponse AccountReadResponse = client.Execute(readRequest);
			var AccountReadResponseToUpdate = (Sugar.Account)AccountReadResponse.Data;

			var request = new SugarRestRequest(RequestType.Update);

			if (String.IsNullOrEmpty(ownerId) || ownerId == Guid.Empty.ToString())
			{
				AccountReadResponseToUpdate.AssignedUserId = AdminSugarUserId;
			}
			else
			{
				AccountReadResponseToUpdate.AssignedUserId = ownerId;
			}

			GetPhoneOfficeAmount();
			AccountReadResponseToUpdate.PhoneFax = "";
			if (AccountEntityCollection.Count == 0)
			{
				AccountReadResponseToUpdate.PhoneOffice = "";
			}
			else if (AccountEntityCollection.Count >= 2)
			{
				AccountReadResponseToUpdate.PhoneOffice = AccountEntityCollection[0].GetTypedColumnValue<string>("Number");
				AccountReadResponseToUpdate.PhoneFax = AccountEntityCollection[1].GetTypedColumnValue<string>("Number");
			}
			else
			{
				AccountReadResponseToUpdate.PhoneOffice = AccountEntityCollection[0].GetTypedColumnValue<string>("Number");
				AccountReadResponseToUpdate.PhoneFax = "";
			}

			AccountReadResponseToUpdate.Name = AccountEntity.GetTypedColumnValue<string>("Name");
			AccountReadResponseToUpdate.AccountType = AccountEntity.GetTypedColumnValue<string>("Type_Name");
			AccountReadResponseToUpdate.Industry = AccountEntity.GetTypedColumnValue<string>("Industry_Name");
			AccountReadResponseToUpdate.Employees = AccountEntity.GetTypedColumnValue<string>("EmployeesNumber_Name");
			AccountReadResponseToUpdate.Website = AccountEntity.GetTypedColumnValue<string>("Web");
			var email = GetAccountEmail();
			AccountReadResponseToUpdate.MainEmailAddress = email;

			request.Parameter = AccountReadResponseToUpdate;

			List<string> selectFields = new List<string>();
			selectFields.Add(nameof(Sugar.Account.AssignedUserId));
			selectFields.Add(nameof(Sugar.Account.Name));
			selectFields.Add(nameof(Sugar.Account.AccountType));
			selectFields.Add(nameof(Sugar.Account.PhoneOffice));
			selectFields.Add(nameof(Sugar.Account.PhoneFax));
			selectFields.Add(nameof(Sugar.Account.Industry));
			selectFields.Add(nameof(Sugar.Account.Employees));
			selectFields.Add(nameof(Sugar.Account.Website));
			selectFields.Add(nameof(Sugar.Account.MainEmailAddress));

			request.Options.SelectFields = selectFields;

			SugarRestResponse response = client.Execute<Sugar.Account>(request);
		}

		private string ReadEntryListResponse()
		{
			string moduleName = "Accounts";

			var client = new SugarRestClient(SugarURL, AdminSugarLogin, AdminSugarPassword);
			var request = new SugarRestRequest(moduleName, RequestType.BulkRead);

			request.Parameter = null;

			List<string> selectFields = new List<string>();
			selectFields.Add(nameof(Sugar.Account.Id));
			request.Options.SelectFields = selectFields;

			var number = GetAccountPhoneNumber();
			var email = GetAccountEmail();
			var queryList = new List<string>();
			if (number != string.Empty)
			{
				queryList.Add($"RIGHT(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(accounts.phone_mobile,' ',''),'+',''),')',''),'(',''),'-',''),'.',''),10) = '{number}'");
			}
			if (email != string.Empty)
			{
				queryList.Add($"contacts.id in (SELECT eabr.bean_id FROM email_addr_bean_rel eabr JOIN email_addresses ea ON (ea.id = eabr.email_address_id) WHERE eabr.deleted=0 and eabr.bean_module like 'accounts' and ea.email_address LIKE '{email}')");
			}
			if (queryList.Count == 0)
			{
				return string.Empty;
			}
			request.Options.Query = String.Join(" OR ", queryList);

			request.Options.MaxResult = 1;

			SugarRestResponse response = client.Execute(request);
			return GetSugarAccountId(response);
		}

		private string GetAccountPhoneNumber()
		{
			var number = AccountEntity.GetTypedColumnValue<string>("Phone");
			if (number == string.Empty)
			{
				return string.Empty;
			}
			number = Regex.Replace(number, @"\D+", "");
			if (number.Length < 10)
			{
				return string.Empty;
			}
			return number.Substring(number.Length - 10, 10);
		}
	}
}