using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(EXCEL2.Startup))]
namespace EXCEL2
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
