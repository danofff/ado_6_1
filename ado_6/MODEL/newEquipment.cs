namespace ado_6.MODEL
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("newEquipment")]
    public partial class newEquipment
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public newEquipment()
        {
            TrackEvaluation = new HashSet<TrackEvaluation>();
        }

        [Key]
        public int intEquipmentID { get; set; }

        [StringLength(50)]
        public string intGarageRoom { get; set; }

        public int intManufacturerID { get; set; }

        public int intModelID { get; set; }

        [StringLength(4)]
        public string strManufYear { get; set; }

        public int intSNPrefixID { get; set; }

        [StringLength(20)]
        public string strSerialNo { get; set; }

        public int intEquipmentTypeID { get; set; }

        public int intSMCSFamilyID { get; set; }

        public int intSizeID { get; set; }

        [Column(TypeName = "date")]
        public DateTime? CreateDate { get; set; }

        public double intMetered { get; set; }

        [Column(TypeName = "date")]
        public DateTime LastDate { get; set; }

        public double intLastMetered { get; set; }

        public double intTotalMetered { get; set; }

        public int? intlimitDay { get; set; }

        public int? intlimitWeek { get; set; }

        public bool bitActive { get; set; }

        public bool bitPreservation { get; set; }

        public bool bitMeter { get; set; }

        public bool bitKTG { get; set; }

        public bool isDelete { get; set; }

        public int intLocationId { get; set; }

        [StringLength(1050)]
        public string strDescription { get; set; }

        public double? floatCostTires { get; set; }

        public int? intCostTiresCurrency { get; set; }

        public double? floatAverageDivergence { get; set; }

        public double? floatFullPrice { get; set; }

        public int? intFullPriceCurrency { get; set; }

        [Column(TypeName = "date")]
        public DateTime? dateStartUpDate { get; set; }

        public int? intServiceLife { get; set; }

        public int? intHoweverOddsAcceleration { get; set; }

        public bool bitMethodAmortization { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TrackEvaluation> TrackEvaluation { get; set; }

        public virtual TablesModel TablesModel { get; set; }
    }
}
