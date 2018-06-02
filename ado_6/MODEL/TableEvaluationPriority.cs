namespace ado_6.MODEL
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TableEvaluationPriority")]
    public partial class TableEvaluationPriority
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TableEvaluationPriority()
        {
            TrackEvaluation = new HashSet<TrackEvaluation>();
        }

        [Key]
        public int intPriorityId { get; set; }

        [StringLength(250)]
        public string strPriorityName { get; set; }

        public int? intLocationId { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TrackEvaluation> TrackEvaluation { get; set; }
    }
}
