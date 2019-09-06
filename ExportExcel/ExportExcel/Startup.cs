using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExportExcel.Startup))]
namespace ExportExcel
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
