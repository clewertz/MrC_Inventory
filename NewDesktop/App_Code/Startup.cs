using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(NewSite.Startup))]
namespace NewSite
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
