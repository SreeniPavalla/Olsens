using Olsens.Plugins.DirectSMSService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Olsens.Plugins.Common
{
    public class SMSData
    {
        private string _message = string.Empty;
        private string _mobilenumber = string.Empty;
        private string _senderId = string.Empty;
        private string _login = string.Empty;
        private string _password = string.Empty;





        public string Message
        {
            get { return _message; }
            set { _message = value; }
        }
        public string MobileNumber
        {
            get { return _mobilenumber; }
            set { _mobilenumber = value; }
        }
        public string senderId
        {
            get { return _senderId; }
            set { _senderId = value; }
        }
        public SMSData(string login, string password)
        {
            _login = login;
            _password = password;



        }
        public string Send()
        {
            try
            {
                SmsGatewayService service = new SmsGatewayService();
                string connectionid = service.connect(_login, _password);

                var mess = new BrandedSmsMessage();
                mess.message = this.Message;
                mess.mobiles = new string[] { this.MobileNumber };
                mess.senderId = this.senderId;

                service.sendBrandedSmsMessage(connectionid, mess);
                return "OK";
            }
            catch (Exception ex) { return "Error: " + ex.Message; }
        }
    }
}
