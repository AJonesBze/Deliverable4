using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class Account
    {
        public Account(string username, string first_name, string last_name, string password, string email, string role_type, string department_name, string status)
        {
            this.username = username ?? throw new ArgumentNullException(nameof(username));
            this.first_name = first_name ?? throw new ArgumentNullException(nameof(first_name));
            this.last_name = last_name ?? throw new ArgumentNullException(nameof(last_name));
            this.password = password ?? throw new ArgumentNullException(nameof(password));
            this.email = email ?? throw new ArgumentNullException(nameof(email));
            this.role_type = role_type ?? throw new ArgumentNullException(nameof(role_type));
            this.department_name = department_name ?? throw new ArgumentNullException(nameof(department_name));
            this.status = status ?? throw new ArgumentNullException(nameof(status));
        }

        [DisplayName("Account Username")]
        public string username { get; set; }

        [DisplayName("User First Name")]
        public string first_name { get; set; }

        [DisplayName("User Last Name")]
        public string last_name { get; set; }

        [DisplayName("Account Password")]
        public string password { get; set; }

        [DisplayName("Account Email Address")]
        public string email { get; set; }

        [DisplayName("Account Role Type")]
        public string role_type { get; set; }

        [DisplayName("Department Name")]
        public string department_name { get; set; }

        [DisplayName("Account Status")]
        public string status { get; set; }

    }
}
