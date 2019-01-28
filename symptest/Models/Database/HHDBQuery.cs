using Dapper;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class HHDBQuery
{
        private static string ServerUsername = "";
        private static string ServerPassword = "";
        private static string ConnectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=master;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        // get single object; based off of https://www.davepaquette.com/archive/2018/01/22/loading-an-object-graph-with-dapper.aspx 
        //        [HttpGet("{id}")]
        //        public async Task<Aircraft> Get(int id)
        //        {
        //            Aircraft aircraft;
        //            using (var connection = new SqlConnection(ConnectionString))
        //            {
        //                await connection.OpenAsync();
        //                var query = @"
        //SELECT 
        //       Id
        //      ,Manufacturer
        //      ,Model
        //      ,RegistrationNumber
        //      ,FirstClassCapacity
        //      ,RegularClassCapacity
        //      ,CrewCapacity
        //      ,ManufactureDate
        //      ,NumberOfEngines
        //      ,EmptyWeight
        //      ,MaxTakeoffWeight
        //  FROM Aircraft WHERE Id = @Id";

        //                aircraft = await connection.QuerySingleAsync<Aircraft>(query, new { Id = id });
        //            }
        //            return aircraft;
        //        }



        // get group of objects; based off of https://www.davepaquette.com/archive/2018/01/22/loading-an-object-graph-with-dapper.aspx
        [HttpGet]
        public async Task<IEnumerable<Gender>> Get()
        {
            IEnumerable<Gender> gender;

            using (var connection = new SqlConnection(ConnectionString))
            {
                await connection.OpenAsync();
                var query = @"SELECT Gender FROM Gender";
                gender = await connection.QueryAsync<Gender>(query);
            }
            return gender;
        }
    }
}
