using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ReleaseNotes.Startup))]
namespace ReleaseNotes
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
