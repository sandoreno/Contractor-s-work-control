namespace ContractorsWorkAPI.Services
{
    public interface IStorageService
    {
        // public async Task<bool> SafeFiles(IFormFile file);
        Task<bool> SafeFiles(IFormFile file);
    }
}
