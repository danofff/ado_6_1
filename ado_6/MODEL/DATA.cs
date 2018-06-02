namespace ado_6.MODEL
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class DATA : DbContext
    {
        public DATA()
            : base("name=DATA")
        {
        }

        public virtual DbSet<newEquipment> newEquipment { get; set; }
        public virtual DbSet<TableEvaluationPriority> TableEvaluationPriority { get; set; }
        public virtual DbSet<TableEvaluationSysStatus> TableEvaluationSysStatus { get; set; }
        public virtual DbSet<TableExchangeRate> TableExchangeRate { get; set; }
        public virtual DbSet<TablesCurrency> TablesCurrency { get; set; }
        public virtual DbSet<TablesModel> TablesModel { get; set; }
        public virtual DbSet<TrackEvaluation> TrackEvaluation { get; set; }
        public virtual DbSet<TrackEvaluationPart> TrackEvaluationPart { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<TableEvaluationPriority>()
                .HasMany(e => e.TrackEvaluation)
                .WithOptional(e => e.TableEvaluationPriority)
                .HasForeignKey(e => e.intPriority);

            modelBuilder.Entity<TablesModel>()
                .HasMany(e => e.newEquipment)
                .WithRequired(e => e.TablesModel)
                .WillCascadeOnDelete(false);
        }
    }
}
