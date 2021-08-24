using Microsoft.Crm.Sdk.Messages;

namespace Olsens.Plugins.Common
{
    public static class Constants
    {
        public enum LogMode
        {
            Debug = 1,
            Error = 2
        };
        public enum AccountType { PrePaid = 1, AtNeed = 2, Completed = 3 }
        public enum CelebrantFees
        {
            Cheque = 1,
            EFT = 2,
            Cash = 3,
            Account = 4,
            //NA = 4,
            FamilyToPay = 5,
            NoPayment = 6
        }
        public enum FlowerStatus
        {
            NotRequired = 1,
            FromFamily = 2,
            FromHouse = 3
        }
        public enum FuneralType
        {
            Cremation = 1,
            Burial = 2,
            Unknown = 3
        }
        public enum TransferFrom
        {
            PlaceOfDeath = 1,
            PlaceOfResidence = 2,
            Other = 3,
            Coroner = 4
        }
        public enum PrePaidStatus
        {
            InActive = 1,
            Active = 2,
            AtNeed = 3
        }
        public enum FuneralStatus
        {
            Quote = 1,
            PreNeed = 2,
            AtNeed = 3,
            PrePaid = 4
        }
        public enum OOSStyle
        {
            A = 1,
            B = 2,
            C = 3
        }
        public enum EmbalmingType
        {
            NotRequired = 1,
            Full = 2,
            Partial = 3
        }
        public enum ChildLifeStatus
        {
            Alive = 1,
            Deceased = 2,
            StillBorn = 3,
            Unknown=4
        }
    }
}
