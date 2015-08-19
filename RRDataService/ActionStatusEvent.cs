using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReconRunner.Model
{
    public enum RequestState
    {
        Unset,                  // Request status has not been set yet
        Succeeded,              // The request succeeded and no issues were found
        CompletedWithWarning,   // The request completed but with one or more cautionary messages
        Error,                  // An error occured, but processing should continue
        FatalError,             // The request did not complete and any further processing should stop
        Information             // This is an informational message and not a status update
    }

    public delegate void ActionStatusEventHandler(object sender, ActionStatusEventArgs e);

    /// <summary>
    /// A simple event class that may be used to publish action status information.
    /// Optionally setting IsDetail to true means that the message is considered a detailed message
    /// that the consumer may only want to show if the user has requested debugging or detailed
    /// info rather than just summary status messages.
    /// </summary>
    public class ActionStatusEventArgs: EventArgs
    {
        public RequestState State { get; set; }
        public string Message { get; set; }
        public DateTime Time { get; set; }
        public bool IsDetail { get; set; }

        public ActionStatusEventArgs(RequestState state, string msg)
        {
            State = state;
            Message = msg;
            Time = DateTime.Now;
            IsDetail = false;
        }

        public ActionStatusEventArgs(RequestState state, string msg, bool detailStatus)
        {
            State = state;
            Message = msg;
            Time = DateTime.Now;
            IsDetail = detailStatus;
        }
    }
}
