using System.Activities;
using System.ComponentModel;





namespace Kali_Mail
{
    public class Send_Mail : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter recipient mail")]
        [DisplayName("To")]
        public InArgument<string> To { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter mailbody")]
        [DisplayName("Body")]
        public InArgument<string> Body { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter subject")]
        [DisplayName("Subject")]
        public InArgument<string> Subject { get; set; }

        [Category("Input")]
        [Description("Is the body HTML?")]
        [DisplayName("Is HTML")]
        public InArgument<bool> IsHtml { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string recipient = To.Get(context);
            string mailBody = Body.Get(context);
            string mailSubject = Subject.Get(context);
            bool isHtml = IsHtml.Get(context);


        }
    }
}
