using Microsoft.EntityFrameworkCore.Migrations;

namespace Json.Da.Migrations
{
    public partial class First : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Disciplines",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Name = table.Column<string>(nullable: true),
                    Competencies = table.Column<string>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Disciplines", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Employees",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Surname = table.Column<string>(nullable: true),
                    Name = table.Column<string>(nullable: true),
                    Fathername = table.Column<string>(nullable: true),
                    Position = table.Column<string>(nullable: true),
                    Rank = table.Column<string>(nullable: true),
                    Rate = table.Column<double>(nullable: false),
                    Chair = table.Column<string>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Employees", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Syllabuses",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    SubjectName = table.Column<string>(nullable: true),
                    PredmetId = table.Column<int>(nullable: true),
                    Year = table.Column<int>(nullable: false),
                    Direction = table.Column<string>(nullable: true),
                    Profile = table.Column<string>(nullable: true),
                    Semester = table.Column<int>(nullable: false),
                    CreditUnits = table.Column<int>(nullable: false),
                    Hours = table.Column<string>(nullable: true),
                    CourseWork = table.Column<string>(nullable: true),
                    SumIndependentWork = table.Column<string>(nullable: true),
                    InteractiveWatch = table.Column<string>(nullable: true),
                    Test = table.Column<bool>(nullable: false),
                    Exam = table.Column<bool>(nullable: false),
                    Lectures = table.Column<int>(nullable: false),
                    LaboratoryExercises = table.Column<int>(nullable: false),
                    Workshops = table.Column<int>(nullable: false),
                    TypesOfLessons = table.Column<string>(nullable: true),
                    AuditoryLessons = table.Column<double>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Syllabuses", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Syllabuses_Disciplines_PredmetId",
                        column: x => x.PredmetId,
                        principalTable: "Disciplines",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Syllabuses_PredmetId",
                table: "Syllabuses",
                column: "PredmetId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Employees");

            migrationBuilder.DropTable(
                name: "Syllabuses");

            migrationBuilder.DropTable(
                name: "Disciplines");
        }
    }
}
