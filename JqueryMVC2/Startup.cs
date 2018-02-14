using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(JqueryMVC2.Startup))]
namespace JqueryMVC2
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
