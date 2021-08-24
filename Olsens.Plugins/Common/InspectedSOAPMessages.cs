using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.ServiceModel.Dispatcher;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Common
{
    public class InspectedSOAPMessages
    {
        public string Request { get; set; }

        public string Response { get; set; }
    }



    namespace MBSGuru
    {
        /// <summary>
        /// Allows capturing of raw SOAP Messages
        /// </summary>
        public class CapturingEndpointBehavior : IEndpointBehavior
        {
            /// <summary>
            /// Holds the messages
            /// </summary>
            public InspectedSOAPMessages SoapMessages { get; set; }

            public CapturingEndpointBehavior(InspectedSOAPMessages soapMessages)
            {
                this.SoapMessages = soapMessages;
            }

            /// <summary>
            /// Required by IEndpointBehavior
            /// </summary>
            /// <param name="endpoint"></param>
            /// <param name="bindingParameters"></param>
            public void AddBindingParameters(ServiceEndpoint endpoint, BindingParameterCollection bindingParameters) { return; }

            /// <summary>
            /// Required by IEndpointBehavior
            /// </summary>
            /// <param name="endpoint"></param>
            /// <param name="clientRuntime"></param>
            public void ApplyClientBehavior(ServiceEndpoint endpoint, ClientRuntime clientRuntime)
            {
                clientRuntime.MessageInspectors.Add(new CapturingMessageInspector(this.SoapMessages));
            }

            /// <summary>
            /// Required by IEndpointBehavior
            /// </summary>
            /// <param name="endpoint"></param>
            /// <param name="endpointDispatcher"></param>
            public void ApplyDispatchBehavior(ServiceEndpoint endpoint, EndpointDispatcher endpointDispatcher) { return; }

            /// <summary>
            /// Required by IEndpointBehavior
            /// </summary>
            /// <param name="endpoint"></param>
            public void Validate(ServiceEndpoint endpoint) { return; }
        }

        /// <summary>
        /// Actual inspector that captures the messages
        /// </summary>
        public class CapturingMessageInspector : IClientMessageInspector
        {
            /// <summary>
            /// Holds the messages
            /// </summary>
            public InspectedSOAPMessages SoapMessages { get; set; }

            public CapturingMessageInspector(InspectedSOAPMessages soapMessages)
            {
                this.SoapMessages = soapMessages;
            }

            /// <summary>
            /// Called after the web service call completes.  Allows capturing of raw response.
            /// </summary>
            /// <param name="reply"></param>
            /// <param name="correlationState"></param>
            public void AfterReceiveReply(ref System.ServiceModel.Channels.Message reply, object correlationState)
            {
                this.SoapMessages.Response = reply.ToString();
            }

            /// <summary>
            /// Called before the web service is invoked.  Allows capturing of raw request.
            /// </summary>
            /// <param name="request"></param>
            /// <param name="channel"></param>
            /// <returns></returns>
            public object BeforeSendRequest(ref System.ServiceModel.Channels.Message request, IClientChannel channel)
            {
                this.SoapMessages.Request = request.ToString();
                return null;
            }
        }
    }
}
