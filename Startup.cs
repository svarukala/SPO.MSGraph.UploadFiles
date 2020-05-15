using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(MicrosoftGraphFilesUpload.Startup))]

namespace MicrosoftGraphFilesUpload
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}