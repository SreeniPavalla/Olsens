using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Model
{
    public class CustomerPaymentItems
    {
        public List<CustomerPayment> Items { get; set; }
    }
    public class InvoiceItems
    {
        public List<ItemInvoice> Items { get; set; }
    }
    public class ProductItems
    {
        public List<Item> Items { get; set; }
    }
    public class JobItems
    {
        public List<Job> Items { get; set; }
    }

    public class EmployeeItems
    {
        public List<Employee> Items { get; set; }
    }

    public class CustomerItems
    {
        public List<Customer> Items { get; set; }
    }

    public class TaxCodeItems
    {
        public List<TaxCode> Items { get; set; }
    }

    public class TaxCode
    {
        public Guid UID { get; set; }
        public string Code { get; set; }
        public string Description { get; set; }
    }

    public class FreightTaxCode
    {
        public Guid UID { get; set; }
        public string Code { get; set; }
        public string Description { get; set; }
    }



    public class Employee
    {
        public Guid UID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public bool IsIndividual { get; set; }
        public string DisplayID { get; set; }
        public bool IsActive { get; set; }
        public List<Address> Addresses { get; set; }
    }

    public class Customer
    {
        public Guid UID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public bool IsIndividual { get; set; }
        public string DisplayID { get; set; }
        public bool IsActive { get; set; }
        public List<Address> Addresses { get; set; }
        public string Notes { get; set; }
        public SellingDetails SellingDetails { get; set; }
    }

    public class Address
    {
        public int Location { get; set; }
        public string Street { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PostCode { get; set; }
        public string Country { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string Phone3 { get; set; }
        public string Fax { get; set; }
        public string Email { get; set; }
        public string Website { get; set; }
        public string ContactName { get; set; }
        public string Salutation { get; set; }
    }

    public class SellingDetails
    {
        public TaxCode TaxCode { get; set; }
        public FreightTaxCode FreightTaxCode { get; set; }
        public Employee SalesPerson { get; set; }
    }

    public class Job
    {
        public Guid UID { get; set; }
        public string Number { get; set; }
        public bool IsHeader { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public Customer LinkedCustomer { get; set; }
    }

    public class ItemInvoice
    {
        public Guid UID { get; set; }
        public Customer Customer { get; set; }
        public Employee Salesperson { get; set; }
        public InvoiceTerms Terms { get; set; }
        public string Date { get; set; }
        public List<ItemInvoiceLine> Lines { get; set; }
        public string CustomerPurchaseOrderNumber { get; set; }
        public string Number { get; set; }
        public string RowVersion { get; set; }

    }
    public class InvoiceTerms
    {
        //
        // Summary:
        //     % monthly charge for late payment.
        public double MonthlyChargeForLatePayment { get; set; }
        //
        // Summary:
        //     The date the discount (if exists) will expire.
        //
        // Remarks:
        //     Available from 2013.5 (cloud), 2014.1 (desktop)
        public DateTime? DiscountExpiryDate { get; set; }
        //
        // Summary:
        //     The discount applicable if the amount if paid before the discount expiry date
        //
        // Remarks:
        //     Available from 2013.5 (cloud), 2014.1 (desktop)
        public decimal? Discount { get; set; }
        //
        // Summary:
        //     The discount applicable in foreign currency if the amount if paid before the
        //     discount expiry date
        public decimal? DiscountForeign { get; set; }
        //
        // Summary:
        //     Date the invoice balance is due.
        //
        // Remarks:
        //     Available from 2013.5 (cloud), 2014.1 (desktop)
        public DateTime? DueDate { get; set; }
        //
        // Summary:
        //     Finance Charge amount applicable to the invoice.
        //
        // Remarks:
        //     Available from 2013.5 (cloud), 2014.1 (desktop)
        public decimal? FinanceCharge { get; set; }
        //
        // Summary:
        //     Finance Charge amount in foreign currency applicable to the invoice.
        public decimal? FinanceChargeForeign { get; set; }
    }

    public class ItemInvoiceLine
    {

        //
        // Summary:
        //     The quantity of goods shipped.
        public decimal ShipQuantity { get; set; }
        //
        // Summary:
        //     Unit price assigned to the item.
        public decimal UnitPrice { get; set; }
        //
        // Summary:
        //     Unit price assigned to the item in foreign currency.
        public decimal? UnitPriceForeign { get; set; }
        //
        // Summary:
        //     Discount rate applicable to the line of the sale invoice.
        public decimal DiscountPercent { get; set; }
        //
        // Summary:
        //     Dollar amount posted to the Asset account setup on an item using 'I Inventory'
        public decimal CostOfGoodsSold { get; set; }
        //
        // Summary:
        //     Item for the invoice line
        public Item Item { get; set; }
        //
        // Summary:
        //     Location of the purchase item bill line
        public Location Location { get; set; }
        //
        // Summary:
        //     Unit of Measure
        public string UnitOfMeasure { get; set; }
        //
        // Summary:
        //     Unit Count
        public decimal? UnitCount { get; set; }
        public decimal Total { get; set; }

        //
        // Summary:
        //     Item for the invoice line
        public Account Account { get; set; }
        public TaxCode TaxCode { get; set; }
        public Job Job { get; set; }
        public string Description { get; set; }
        public string RowVersion { get; set; }

    }

    public class Item
    {
        public Guid UID { get; set; }
        public string Number { get; set; }
        public string Name { get; set; }
        public bool IsActive { get; set; }
    }

    public class Location
    {
        public Guid UID { get; set; }
        public string Identifier { get; set; }
        public string Name { get; set; }
        public string URI { get; set; }
    }

    public class Account
    {
        public Guid UID { get; set; }
        public string Name { get; set; }
        public string DisplayID { get; set; }
        public string URI { get; set; }
    }

    public class CustomerPayment
    {
        public Guid UID { get; set; }
        public Account Account { get; set; }
        public Customer Customer { get; set; }
        public string ReceiptNumber { get; set; }
        public DateTime Date { get; set; }
        public decimal AmountReceived { get; set; }
        public List<PaymentInvoice> Invoices { get; set; }
    }

    public class PaymentInvoice
    {
        public Guid UID { get; set; }
        public string RowId { get; set; }
        public string Number { get; set; }
        public decimal AmountApplied { get; set; }
        public string ReceiptNumber { get; set; }

    }

    public class Response
    {
        public string Message { get; set; }
        public string InvoiceNumber { get; set; }
        public bool Success { get; set; }
        public string access_token { get; set; }
        public string refresh_token { get; set; }
    }

    public class Errors
    {
        public string Message { get; set; }
        public string Name { get; set; }
        public string AdditionalDetails { get; set; }
    }
    public class APIErrorResponse
    {
        public List<Errors> Errors { get; set; }
    }

    public class BDMFuneralData
    {
        public Guid opportunityid { get; set; }
        public int BDM_Status { get; set; }
        public string Error { get; set; }
        public Guid XMLId { get; set; }
        public string ValidationError { get; set; }
        public string Notificationid { get; set; }
        public string Notificationmessage { get; set; }
        public int Notificationstatus { get; set; }
        public string Applicationid { get; set; }
        public string Applicationmessage { get; set; }
        public string Applicationstatus { get; set; }


    }

    public class LeadDetails
    {
        public string firstname { get; set; }

        public string lastname { get; set; }
       
        public string subject { get; set; }

        public string email { get; set; }

        public string telephone { get; set; }

        public string postcode { get; set; }

        public string description { get; set; }

    }
    public class CompanyFile
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string LibraryPath { get; set; }
        public string ProductVersion { get; set; }
        public string Uri { get; set; }
    }
    public enum BDMStatus { ReadyToSubmit = 1, Submitted = 2, RegistrationError = 3, ValidationError = 4 };
}
