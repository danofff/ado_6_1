namespace ado_6.MODEL
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TablesCurrency")]
    public partial class TablesCurrency
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TablesCurrency()
        {
            TableExchangeRate = new HashSet<TableExchangeRate>();
        }

        [Key]
        public int intCurrencyId { get; set; }

        [StringLength(50)]
        public string strCurrencyName { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TableExchangeRate> TableExchangeRate { get; set; }
    }
}
