namespace NPOItest.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Account")]
    public partial class Account
    {
        public int Id { get; set; }

        [Required]
        [StringLength(255)]
        public string Username { get; set; }

        [Required]
        [StringLength(256)]
        public string Password { get; set; }

        [Required]
        [StringLength(256)]
        public string Name { get; set; }

        [Required]
        [StringLength(256)]
        public string Email { get; set; }

        [Required]
        [StringLength(256)]
        public string Sex { get; set; }

        [Required]
        [StringLength(256)]
        public string Company { get; set; }

        [Required]
        [StringLength(256)]
        public string Position { get; set; }

        [Required]
        [StringLength(256)]
        public string Phone { get; set; }
    }
}
