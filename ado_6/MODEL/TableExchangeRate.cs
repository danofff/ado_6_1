namespace ado_6.MODEL
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TableExchangeRate")]
    public partial class TableExchangeRate
    {
        [Key]
        public int intExchangeRateId { get; set; }

        public int? intCurrencyId { get; set; }

        public double? fRate { get; set; }

        [Column(TypeName = "date")]
        public DateTime? dDateRate { get; set; }

        public virtual TablesCurrency TablesCurrency { get; set; }
    }
}
