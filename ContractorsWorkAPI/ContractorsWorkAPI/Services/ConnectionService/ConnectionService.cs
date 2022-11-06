using ContractorsWorkAPI.Services.Impl;

namespace ContractorsWorkAPI.Services.ConnectionService
{
    public static class ConnectionService
    {
        /// <summary>
        /// прописывается создание сервисов
        /// </summary>
        /// <param name="builder"></param>
        public static void ConnectService(WebApplicationBuilder builder) 
        {
            builder.Services.AddTransient<IStorageService, StorageService>();
        }
    }
}
