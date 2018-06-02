namespace ado_6.MODEL
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TablesModel")]
    public partial class TablesModel
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TablesModel()
        {
            newEquipment = new HashSet<newEquipment>();
        }

        [Key]
        public int intModelID { get; set; }

        [StringLength(10)]
        public string strName { get; set; }

        public int? intManufacturerID { get; set; }

        public int? intSMCSFamilyID { get; set; }

        [StringLength(250)]
        public string strImage { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<newEquipment> newEquipment { get; set; }
    }
}
