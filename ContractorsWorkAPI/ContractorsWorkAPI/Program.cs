using ContractorsWorkAPI.Data;
using ContractorsWorkAPI.Services;
using ContractorsWorkAPI.Services.ConnectionService;
using ContractorsWorkAPI.Services.Impl;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// ���������� ��������
ConnectionService.ConnectService(builder);  

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();


//����������� � ��
builder.Services.AddDbContext<DataContext>(op =>
{
    // ����� ��������� ����������� � ��
    op.UseNpgsql(builder.Configuration.GetConnectionString("EmployeeAppCon"));
});


var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

AppContext.SetSwitch("Npgsql.EnableLegacyTimestampBehavior", true);

app.Run();
