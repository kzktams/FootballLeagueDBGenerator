using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Formats.Asn1;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Generate
{
    class Program
    {
        static void Main(string[] args)
        {
            var random = new Random();
            var seasons = GenerateSeasons();
            var leagues = GenerateLeagues();
            var teams = LoadTeamsFromExcel("foci.xlsx");
            var positions = LoadPositionsFromExcel("pozi.xlsx");
            var players = LoadPlayersFromExcel("players.xlsx", positions);
            var stadiums = GenerateStadiums(teams, 200);
            var coaches = LoadCoachesFromExcel("coaches.xlsx");
            var referee = LoadRefereesFromExcel("referee.xlsx");
            var coachAssignment = GenerateCoachAssignments(coaches,teams,seasons,random);
            var playerAssignments = GeneratePlayerAssignments(players, teams, seasons, random);
            var matches = GenerateMatches(teams,seasons,stadiums,referee,random);
            var goals = GenerateGoals(matches, playerAssignments, random);
            SaveGoalsToCsv(goals, "goals.csv");
            AssignStadiumsToTeams(teams, stadiums);
            ExportDataToExcel(seasons, leagues, teams, stadiums,players,positions,coaches,referee, coachAssignment, playerAssignments,matches, "GeneratedData.xlsx");
        }
        public static void SaveGoalsToCsv(List<Goal> goals, string filePath)
        {
            using (var writer = new StreamWriter(filePath))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.WriteRecords(goals);
            }
        }
        static List<Goal> GenerateGoals(List<Match> matches, List<PlayerAssignment> playerAssignments, Random random)
        {
            var goals = new List<Goal>();
            int goalId = 1;

            foreach (var match in matches)
            {
                var homePlayers = playerAssignments.Where(pa => pa.TeamId == match.HomeTeamId && pa.SeasonId == match.SeasonId).Select(pa => pa.PlayerId).ToList();
                var awayPlayers = playerAssignments.Where(pa => pa.TeamId == match.AwayTeamId && pa.SeasonId == match.SeasonId).Select(pa => pa.PlayerId).ToList();

                // Determine goals count for each team
                int homeGoalsCount = random.Next(1, Math.Max(1, homePlayers.Count / 2));  // Ensure there are enough players
                int awayGoalsCount = random.Next(1, Math.Max(1, awayPlayers.Count / 2));  // Ensure there are enough players

                if (match.WinnerTeamId != null)
                {
                    if (match.WinnerTeamId == match.HomeTeamId)
                    {
                        awayGoalsCount = Math.Min(homeGoalsCount - 1, awayPlayers.Count / 2);  // Home team wins
                    }
                    else
                    {
                        homeGoalsCount = Math.Min(awayGoalsCount - 1, homePlayers.Count / 2);  // Away team wins
                    }
                }
                else
                {
                    int minGoals = Math.Min(homeGoalsCount, awayGoalsCount);
                    homeGoalsCount = awayGoalsCount = minGoals;  // Draw
                }

                // Generate goals for both teams
                goals.AddRange(GenerateTeamGoals(match, homePlayers, homeGoalsCount, ref goalId, match.HomeTeamId, match.SeasonId, match.LeagueId, random));
                goals.AddRange(GenerateTeamGoals(match, awayPlayers, awayGoalsCount, ref goalId, match.AwayTeamId, match.SeasonId, match.LeagueId, random));
            }

            return goals;
        }

        static List<Goal> GenerateTeamGoals(Match match, List<int> players, int goalsCount, ref int goalId, int teamId, int seasonId, int leagueId, Random random)
        {
            List<Goal> teamGoals = new List<Goal>();
            if (players.Count < 1) return teamGoals;  // Ha kevesebb, mint 1 játékos van, nem generálunk gólokat

            for (int i = 0; i < goalsCount; i++)
            {
                int scorerIndex = random.Next(players.Count);
                int scorer = players[scorerIndex];

                teamGoals.Add(new Goal
                {
                    GoalId = goalId++,
                    MatchId = match.MatchId,
                    PlayerId = scorer,
                    TeamId = teamId,
                    SeasonId = seasonId,
                    LeagueId = leagueId
                });
            }
            
            return teamGoals;
            
        }








        static DateTime GenerateBirthDate()
        {
            var random = new Random();
            var start = new DateTime(1970, 1, 1);
            var end = new DateTime(2007, 12, 31);
            int range = (end - start).Days;
            return start.AddDays(random.Next(range));
        }

        static List<Match> GenerateMatches(List<Team> teams, List<Season> seasons, List<Stadium> stadiums, List<Referee> referees, Random random)
        {
            var matches = new List<Match>();
            int matchCounter = 1;
            var firstFiveSeasons = seasons.Take(5).ToList();

            foreach (var season in seasons)
            {
                foreach (var leagueId in teams.Select(t => t.LeagueId).Distinct())
                {
                    var leagueTeams = teams.Where(t => t.LeagueId == leagueId).ToList();
                    for (int i = 0; i < leagueTeams.Count; i++)
                    {
                        for (int j = i + 1; j < leagueTeams.Count; j++)
                        {
                            var homeTeam = leagueTeams[i];
                            var awayTeam = leagueTeams[j];

                            // Home match
                            var homeMatch = CreateMatch(matchCounter++, homeTeam, awayTeam, season, stadiums, referees, leagueId, random);
                            matches.Add(homeMatch);

                            // Away match
                            var awayMatch = CreateMatch(matchCounter++, awayTeam, homeTeam, season, stadiums, referees, leagueId, random);
                            matches.Add(awayMatch);
                        }
                    }
                }
            }

            return matches;
        }

        static Match CreateMatch(int matchId, Team homeTeam, Team awayTeam, Season season, List<Stadium> stadiums, List<Referee> referees, int leagueId, Random random)
        {
            return new Match
            {
                MatchId = $"M{matchId}",
                HomeTeamId = homeTeam.TeamId,
                AwayTeamId = awayTeam.TeamId,
                StadiumId = homeTeam.StadiumId, // Assuming each team has a stadium
                RefereeId = referees[random.Next(referees.Count)].RefereeId,
                Date = RandomMatchDate(season, random),
                Attendance = random.Next(5000, 50000),
                SeasonId = season.SeasonId,
                WinnerTeamId = DecideWinner(homeTeam, awayTeam, random),
                LeagueId = leagueId
            };
        }


        static DateTime RandomMatchDate(Season season, Random random)
        {
            var days = (season.EndDate - season.StartDate).Days;
            return season.StartDate.AddDays(random.Next(days));
        }

        static int? DecideWinner(Team homeTeam, Team awayTeam, Random random)
        {
            int result = random.Next(3); // 0 for home win, 1 for away win, 2 for draw
            if (result == 0) return homeTeam.TeamId;
            if (result == 1) return awayTeam.TeamId;
            return null; // Draw
        }


        static List<Season> GenerateSeasons()
        {
            var seasons = new List<Season>();
            int currentYear = DateTime.Now.Year;
            for (int i = 0; i < 20; i++)
            {
                int startYear = currentYear - i - 1;
                int endYear = startYear + 1;
                seasons.Add(new Season
                {
                    SeasonId = i + 1,
                    Name = $"{startYear}/{endYear}",
                    StartDate = new DateTime(startYear, 8, 1),
                    EndDate = new DateTime(endYear, 5, 31)
                });
            }
            return seasons;
        }

        static List<League> GenerateLeagues()
        {
            var leagues = new List<League>();
            for (int i = 1; i <= 8; i++)
            {
                leagues.Add(new League
                {
                    LeagueId = i,
                    Name = $"League {i}"
                });
            }
            return leagues;
        }

        static List<Referee> LoadRefereesFromExcel(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var referees = new List<Referee>();
            var random = new Random();
            int refereeNumber = 1;

            foreach (var row in worksheet.RangeUsed().RowsUsed().Skip(1))
            {
                var referee = new Referee
                {
                    RefereeId = GenerateRefereeId(refereeNumber++),
                    FirstName = row.Cell(1).GetValue<string>(),
                    LastName = row.Cell(2).GetValue<string>(),
                    Nationality = row.Cell(3).GetValue<string>(),
                    Email = row.Cell(4).GetValue<string>(),
                    PhoneNumber = row.Cell(5).GetValue<string>(),
                    BirthDate = GenerateBirthDateForCoach()
                };
                referees.Add(referee);
            }
            return referees;
        }
        static string GenerateRefereeId(int number)
        {
            return $"REF{number.ToString("D4")}";
        }
        static List<Coach> LoadCoachesFromExcel(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var coaches = new List<Coach>();
            var random = new Random();
            int coachNumber = 1;

            foreach (var row in worksheet.RangeUsed().RowsUsed().Skip(1))
            {
                var birthDate = GenerateBirthDateForCoach();
                var coachId = GenerateCoachId(coachNumber++, random);

                coaches.Add(new Coach
                {
                    CoachId = coachId,
                    FirstName = row.Cell(1).GetValue<string>(),
                    LastName = row.Cell(2).GetValue<string>(),
                    Nationality = row.Cell(3).GetValue<string>(),
                    ExperienceLevel = row.Cell(4).GetValue<string>(),
                    CoachingLicenses = row.Cell(5).GetValue<int>(),
                    CoachingStyle = row.Cell(6).GetValue<string>(),
                    TrainingMethods = row.Cell(7).GetValue<string>(),
                    PlayerDevelopmentFocus = row.Cell(8).GetValue<string>(),
                    InjuryManagement = row.Cell(9).GetValue<string>(),
                    TeamSelectionCriteria = row.Cell(10).GetValue<string>(),
                    CommunicationStyle = row.Cell(11).GetValue<string>(),
                    MotivationalTechniques = row.Cell(12).GetValue<string>(),
                    BirthDate = birthDate
                });
            }
            return coaches;
        }



        static DateTime GenerateBirthDateForCoach()
        {
            var random = new Random();
            var start = new DateTime(1940, 1, 1);
            var end = new DateTime(2000, 12, 31);
            int range = (end - start).Days;
            return start.AddDays(random.Next(range));
        }

        static string GenerateCoachId(int number, Random random)
        {
            const string chars = "abcdefghijklmnopqrstuvwxyz";
            var randomChars = new string(Enumerable.Repeat(chars, 4)
                .Select(s => s[random.Next(s.Length)]).ToArray());
            return $"{number}{randomChars}";
        }

        static List<Team> LoadTeamsFromExcel(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var teams = worksheet.RangeUsed().RowsUsed() // Kihagyja a fejlécet
                .Select((row, index) => new Team
                {
                    TeamId = index + 1,
                    Name = row.Cell(1).GetValue<string>(),
                    City = row.Cell(2).GetValue<string>(),
                    EstablishmentYear = new Random().Next(1890, 2001),
                    LeagueId = (index / 30) + 1
                }).ToList();

            return teams;
        }
        static List<Position> LoadPositionsFromExcel(string filePath)
        {
            var positions = new List<Position>();
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            foreach (var row in worksheet.RangeUsed().RowsUsed()) // Kihagyja a fejlécet
            {
                positions.Add(new Position
                {
                    PositionId = row.Cell(1).GetValue<string>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Category = row.Cell(3).GetValue<string>()
                });
            }
            return positions;
        }

        static List<Player> LoadPlayersFromExcel(string filePath, List<Position> positions)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var players = new List<Player>();
            var random = new Random();
            int goalkeeperCount = 0;
            int playerId = 1;

            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                var player = new Player
                {
                    PlayerId = playerId++,
                    FirstName = row.Cell(1).GetValue<string>(),
                    LastName = row.Cell(2).GetValue<string>(),
                    Nationality = row.Cell(3).GetValue<string>(),
                    BirthDate = GenerateBirthDate(),
                    Height = random.Next(160, 211),
                    PositionId = AssignPositionId(positions, ref goalkeeperCount, 4500)
                };
                players.Add(player);
            }
            return players;
        }

        static List<CoachAssignment> GenerateCoachAssignments(List<Coach> coaches, List<Team> teams, List<Season> seasons, Random random)
        {
            var assignments = new List<CoachAssignment>();
            var coachAssignmentsByTeam = new Dictionary<int, string>(); // key: TeamId, value: last CoachId
            var occupiedCoachesPerSeason = new Dictionary<int, HashSet<string>>(); // key: SeasonId, value: HashSet of CoachIds

            foreach (var season in seasons)
            {
                var seasonCoaches = new HashSet<string>(); // Track coaches assigned this season

                foreach (var team in teams)
                {
                    string lastCoachId = coachAssignmentsByTeam.ContainsKey(team.TeamId) ? coachAssignmentsByTeam[team.TeamId] : null;
                    var eligibleCoaches = coaches.Where(c => c.BirthDate.AddYears(20) <= season.StartDate).ToList();

                    string coachId = null;
                    if (lastCoachId != null && random.NextDouble() < 0.7 && !seasonCoaches.Contains(lastCoachId))
                    {
                        coachId = lastCoachId;
                    }
                    else
                    {
                        var newCoachOptions = eligibleCoaches.Where(c => !seasonCoaches.Contains(c.CoachId)).ToList();
                        if (newCoachOptions.Count > 0)
                        {
                            coachId = newCoachOptions[random.Next(newCoachOptions.Count)].CoachId;
                        }
                    }

                    if (coachId != null)
                    {
                        seasonCoaches.Add(coachId); // Mark this coach as occupied for this season
                        var assignment = new CoachAssignment
                        {
                            AssignmentId = $"ASGN{assignments.Count + 1}",
                            CoachId = coachId,
                            TeamId = team.TeamId,
                            SeasonId = season.SeasonId
                        };
                        assignments.Add(assignment);
                        coachAssignmentsByTeam[team.TeamId] = coachId; // Update last coach for the team
                    }
                }
                occupiedCoachesPerSeason[season.SeasonId] = seasonCoaches; // Store occupied coaches for this season
            }

            return assignments;
        }


        static List<PlayerAssignment> GeneratePlayerAssignments(List<Player> players, List<Team> teams, List<Season> seasons, Random random)
        {
            var assignments = new List<PlayerAssignment>();
            var playerAssignmentsBySeason = new Dictionary<int, List<int>>(); // key: SeasonId, value: List of PlayerIds

            foreach (var season in seasons)
            {
                // Gather players who are old enough to play this season
                var eligiblePlayers = players.Where(p => p.BirthDate.AddYears(16) <= season.StartDate).ToList();

                foreach (var team in teams)
                {
                    int teamPlayerCount = random.Next(20, 51); // Randomly decide how many players this team will have this season
                    var teamPlayers = new HashSet<int>(); // To ensure unique players per team per season

                    while (teamPlayers.Count < teamPlayerCount && eligiblePlayers.Count > 0)
                    {
                        var player = eligiblePlayers[random.Next(eligiblePlayers.Count)];
                        if (!playerAssignmentsBySeason.ContainsKey(season.SeasonId) || !playerAssignmentsBySeason[season.SeasonId].Contains(player.PlayerId))
                        {
                            teamPlayers.Add(player.PlayerId);
                            eligiblePlayers.Remove(player); // Remove player from eligible pool
                            if (!playerAssignmentsBySeason.ContainsKey(season.SeasonId))
                            {
                                playerAssignmentsBySeason[season.SeasonId] = new List<int>();
                            }
                            playerAssignmentsBySeason[season.SeasonId].Add(player.PlayerId);

                            assignments.Add(new PlayerAssignment
                            {
                                AssignmentId = $"PASGN{assignments.Count + 1}",
                                PlayerId = player.PlayerId,
                                TeamId = team.TeamId,
                                SeasonId = season.SeasonId
                            });
                        }
                    }
                }
            }

            return assignments;
        }




        static string AssignPositionId(List<Position> positions, ref int goalkeeperCount, int maxGoalkeepers)
        {
            var random = new Random();
            Position position;
            do
            {
                position = positions[random.Next(positions.Count)];
            } while (position.Name == "Goalkeeper" && goalkeeperCount >= maxGoalkeepers);

            if (position.Name == "Goalkeeper")
            {
                goalkeeperCount++;
            }

            return position.PositionId;
        }
        static List<Stadium> GenerateStadiums(List<Team> teams, int count)
        {
            var random = new Random();
            var stadiums = new List<Stadium>();

            for (int i = 1; i <= count; i++)
            {
                var city = teams[random.Next(teams.Count)].City;
                stadiums.Add(new Stadium
                {
                    StadiumId = $"S{i}{random.Next(1000,10000)}", // Előtag hozzáadása az ID-hoz
                    Name = $"Stadium {i}",
                    Capacity = random.Next(5000, 50000),
                    City = city,
                    Address = $"Address {i} in {city}"
                });
            }

            return stadiums;
        }

        static void AssignStadiumsToTeams(List<Team> teams, List<Stadium> stadiums)
        {
            var random = new Random();
            var availableStadiums = new List<Stadium>(stadiums); // Elérhető stadionok másolata

            // Első körben minden stadionhoz hozzárendelünk egy csapatot
            for (int i = 0; i < teams.Count && availableStadiums.Any(); i++)
            {
                var stadium = availableStadiums[0]; // Mindig az első elérhető stadiont vesszük
                teams[i].StadiumId = stadium.StadiumId;
                stadium.City = teams[i].City; // Frissítjük a stadion városát, ha szükséges
                availableStadiums.RemoveAt(0); // Eltávolítjuk a hozzárendelt stadiont az elérhetők listájából
            }

            // Ha több csapat van, mint stadion, a maradék csapatokhoz véletlenszerű stadionok hozzárendelése
            if (availableStadiums.Count == 0 && teams.Count > stadiums.Count)
            {
                for (int i = stadiums.Count; i < teams.Count; i++)
                {
                    var stadium = stadiums[random.Next(stadiums.Count)]; // Véletlenszerű stadion kiválasztása
                    teams[i].StadiumId = stadium.StadiumId;
                }
            }
        }

        static void ExportDataToExcel(
            List<Season> seasons, List<League> leagues ,List<Team> teams, 
            List<Stadium> stadiums, List<Player> players, List<Position> positions,
            List<Coach> coaches, List<Referee> referees,
            List<CoachAssignment> coachAssignments,
            List<PlayerAssignment> playerAssignments,
            List<Match> matches,
            
            string filePath)
        {
            var workbook = new XLWorkbook();
            ExportSeasonsToExcel(workbook, seasons);
            ExportLeaguesToExcel(workbook, leagues);
            ExportTeams(workbook, teams);
            ExportStadiums(workbook, stadiums);
            ExportPositions(workbook, positions);
            ExportPlayers(workbook, players);
            ExportCoachesToExcel(workbook, coaches);
            ExportRefereesToExcel(workbook, referees);
            ExportCoachAssignmentsToExcel(workbook, coachAssignments);
            ExportPlayerAssignmentsToExcel(workbook, playerAssignments);
            ExportMatchesToExcel(workbook, matches);
            
            workbook.SaveAs(filePath);
        }
        static void ExportGoalsToExcel(XLWorkbook workbook, List<Goal> goals)
        {
            var sheet = workbook.AddWorksheet("Goals");
            sheet.Cell(1, 1).Value = "GoalId";
            sheet.Cell(1, 2).Value = "MatchId";
            sheet.Cell(1, 3).Value = "PlayerId";
            sheet.Cell(1, 4).Value = "AssisterId";
            sheet.Cell(1, 5).Value = "TeamId";
            sheet.Cell(1, 6).Value = "SeasonId";
            sheet.Cell(1, 7).Value = "LeagueId";

            int row = 2;
            foreach (var goal in goals)
            {
                sheet.Cell(row, 1).Value = goal.GoalId;
                sheet.Cell(row, 2).Value = goal.MatchId;
                sheet.Cell(row, 3).Value = goal.PlayerId;
                //sheet.Cell(row, 4).Value = goal.AssisterId;
                sheet.Cell(row, 5).Value = goal.TeamId;
                sheet.Cell(row, 6).Value = goal.SeasonId;
                sheet.Cell(row, 7).Value = goal.LeagueId;
                row++;
            }
        }


        static void ExportMatchesToExcel(XLWorkbook workbook, List<Match> matches)
        {
            var sheet = workbook.AddWorksheet("Matches");
            sheet.Cell(1, 1).Value = "MatchId";
            sheet.Cell(1, 2).Value = "HomeTeamId";
            sheet.Cell(1, 3).Value = "AwayTeamId";
            sheet.Cell(1, 4).Value = "StadiumId";
            sheet.Cell(1, 5).Value = "RefereeId";
            sheet.Cell(1, 6).Value = "Date";
            sheet.Cell(1, 7).Value = "Attendance";
            sheet.Cell(1, 8).Value = "SeasonId";
            sheet.Cell(1, 9).Value = "WinnerTeamId";
            sheet.Cell(1, 10).Value = "LeagueId";

            int row = 2;
            foreach (var match in matches)
            {
                sheet.Cell(row, 1).Value = match.MatchId;
                sheet.Cell(row, 2).Value = match.HomeTeamId;
                sheet.Cell(row, 3).Value = match.AwayTeamId;
                sheet.Cell(row, 4).Value = match.StadiumId;
                sheet.Cell(row, 5).Value = match.RefereeId;
                sheet.Cell(row, 6).Value = match.Date.ToString("yyyy-MM-dd");
                sheet.Cell(row, 7).Value = match.Attendance;
                sheet.Cell(row, 8).Value = match.SeasonId;
                sheet.Cell(row, 9).Value = match.WinnerTeamId;
                sheet.Cell(row, 10).Value = match.LeagueId;
                row++;
            }
        }



        static void ExportPlayerAssignmentsToExcel(XLWorkbook workbook, List<PlayerAssignment> assignments)
        {
            var sheet = workbook.AddWorksheet("PlayerAssignments");
            sheet.Cell(1, 1).Value = "AssignmentId";
            sheet.Cell(1, 2).Value = "PlayerId";
            sheet.Cell(1, 3).Value = "TeamId";
            sheet.Cell(1, 4).Value = "SeasonId";

            int row = 2;
            foreach (var assignment in assignments)
            {
                sheet.Cell(row, 1).Value = assignment.AssignmentId;
                sheet.Cell(row, 2).Value = assignment.PlayerId;
                sheet.Cell(row, 3).Value = assignment.TeamId;
                sheet.Cell(row, 4).Value = assignment.SeasonId;
                row++;
            }
        }

        static void ExportCoachAssignmentsToExcel(XLWorkbook workbook, List<CoachAssignment> assignments)
        {
            var sheet = workbook.AddWorksheet("CoachAssignments");
            sheet.Cell(1, 1).Value = "AssignmentId";
            sheet.Cell(1, 2).Value = "CoachId";
            sheet.Cell(1, 3).Value = "TeamId";
            sheet.Cell(1, 4).Value = "SeasonId";
            

            int row = 2;
            foreach (var assignment in assignments)
            {
                sheet.Cell(row, 1).Value = assignment.AssignmentId;
                sheet.Cell(row, 2).Value = assignment.CoachId;
                sheet.Cell(row, 3).Value = assignment.TeamId;
                sheet.Cell(row, 4).Value = assignment.SeasonId;
                
                row++;
            }
        }

        static void ExportRefereesToExcel(XLWorkbook workbook, List<Referee> referees)
        {
            var sheet = workbook.AddWorksheet("Referees");
            sheet.Cell(1, 1).Value = "RefereeId";
            sheet.Cell(1, 2).Value = "FirstName";
            sheet.Cell(1, 3).Value = "LastName";
            sheet.Cell(1, 4).Value = "Nationality";
            sheet.Cell(1, 5).Value = "Email";
            sheet.Cell(1, 6).Value = "PhoneNumber";
            sheet.Cell(1, 7).Value = "BirthDate";

            int row = 2;
            foreach (var referee in referees)
            {
                sheet.Cell(row, 1).Value = referee.RefereeId;
                sheet.Cell(row, 2).Value = referee.FirstName;
                sheet.Cell(row, 3).Value = referee.LastName;
                sheet.Cell(row, 4).Value = referee.Nationality;
                sheet.Cell(row, 5).Value = referee.Email;
                sheet.Cell(row, 6).Value = referee.PhoneNumber;
                sheet.Cell(row, 7).Value = referee.BirthDate.ToString("yyyy-MM-dd");
                row++;
            }
        }

        static void ExportCoachesToExcel(XLWorkbook workbook, List<Coach> coaches)
        {
            var sheet = workbook.AddWorksheet("Coaches");
            sheet.Cell(1, 1).Value = "CoachId";
            sheet.Cell(1, 2).Value = "FirstName";
            sheet.Cell(1, 3).Value = "LastName";
            sheet.Cell(1, 4).Value = "Nationality";
            sheet.Cell(1, 5).Value = "Experience Level";
            sheet.Cell(1, 6).Value = "Coaching Licenses";
            sheet.Cell(1, 7).Value = "Coaching Style";
            sheet.Cell(1, 8).Value = "Training Methods";
            sheet.Cell(1, 9).Value = "Player Development Focus";
            sheet.Cell(1, 10).Value = "Injury Management";
            sheet.Cell(1, 11).Value = "Team Selection Criteria";
            sheet.Cell(1, 12).Value = "Communication Style";
            sheet.Cell(1, 13).Value = "Motivational Techniques";
            sheet.Cell(1, 14).Value = "BirthDate";

            int row = 2;
            foreach (var coach in coaches)
            {
                sheet.Cell(row, 1).Value = coach.CoachId;
                sheet.Cell(row, 2).Value = coach.FirstName;
                sheet.Cell(row, 3).Value = coach.LastName;
                sheet.Cell(row, 4).Value = coach.Nationality;
                sheet.Cell(row, 5).Value = coach.ExperienceLevel;
                sheet.Cell(row, 6).Value = coach.CoachingLicenses;
                sheet.Cell(row, 7).Value = coach.CoachingStyle;
                sheet.Cell(row, 8).Value = coach.TrainingMethods;
                sheet.Cell(row, 9).Value = coach.PlayerDevelopmentFocus;
                sheet.Cell(row, 10).Value = coach.InjuryManagement;
                sheet.Cell(row, 11).Value = coach.TeamSelectionCriteria;
                sheet.Cell(row, 12).Value = coach.CommunicationStyle;
                sheet.Cell(row, 13).Value = coach.MotivationalTechniques;
                sheet.Cell(row, 14).Value = coach.BirthDate.ToString("yyyy-MM-dd");
                row++;
            }
        }



        static void ExportSeasonsToExcel(XLWorkbook workbook, List<Season> seasons)
        {
            var sheet = workbook.AddWorksheet("Seasons");
            sheet.Cell(1, 1).Value = "SeasonId";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "StartDate";
            sheet.Cell(1, 4).Value = "EndDate";

            int row = 2;
            foreach (var season in seasons)
            {
                sheet.Cell(row, 1).Value = season.SeasonId;
                sheet.Cell(row, 2).Value = season.Name;
                sheet.Cell(row, 3).Value = season.StartDate.ToString("yyyy-MM-dd");
                sheet.Cell(row, 4).Value = season.EndDate.ToString("yyyy-MM-dd");
                row++;
            }
        }

        static void ExportLeaguesToExcel(XLWorkbook workbook, List<League> leagues)
        {
            var sheet = workbook.AddWorksheet("Leagues");
            sheet.Cell(1, 1).Value = "LeagueId";
            sheet.Cell(1, 2).Value = "Name";

            int row = 2;
            foreach (var league in leagues)
            {
                sheet.Cell(row, 1).Value = league.LeagueId;
                sheet.Cell(row, 2).Value = league.Name;
                row++;
            }
        }

        static void ExportPositions(XLWorkbook workbook, List<Position> positions)
        {
            var positionsSheet = workbook.Worksheets.Add("Positions");

            positionsSheet.Cell(1, 1).Value = "PositionId";
            positionsSheet.Cell(1, 2).Value = "Name";
            positionsSheet.Cell(1, 3).Value = "Category";
            int posRow = 2;
            foreach (var position in positions)
            {
                positionsSheet.Cell(posRow, 1).Value = position.PositionId;
                positionsSheet.Cell(posRow, 2).Value = position.Name;
                positionsSheet.Cell(posRow, 3).Value = position.Category;
                posRow++;

            }
        }

        static void ExportPlayers(XLWorkbook workbook, List<Player> players) 
        {
            var playersSheet = workbook.Worksheets.Add("Players");
            playersSheet.Cell(1, 1).Value = "PlayerId";
            playersSheet.Cell(1, 2).Value = "FirstName";
            playersSheet.Cell(1, 3).Value = "LastName";
            playersSheet.Cell(1, 4).Value = "Nationality";
            playersSheet.Cell(1, 5).Value = "BirthDate";
            playersSheet.Cell(1, 6).Value = "Height";
            playersSheet.Cell(1, 7).Value = "PositionId";
            int playerRow = 2;
            foreach (var player in players)
            {
                playersSheet.Cell(playerRow, 1).Value = player.PlayerId;
                playersSheet.Cell(playerRow, 2).Value = player.FirstName;
                playersSheet.Cell(playerRow, 3).Value = player.LastName;
                playersSheet.Cell(playerRow, 4).Value = player.Nationality;
                playersSheet.Cell(playerRow, 5).Value = player.BirthDate.ToString("yyyy-MM-dd");
                playersSheet.Cell(playerRow, 6).Value = player.Height;
                playersSheet.Cell(playerRow, 7).Value = player.PositionId;
                playerRow++;
            }
        }

        // ExportTeams és ExportStadiums metódusok implementációja
        static void ExportTeams(XLWorkbook workbook, List<Team> teams)
        {
            var sheet = workbook.AddWorksheet("Teams");
            sheet.Cell(1, 1).Value = "TeamId";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "City";
            sheet.Cell(1, 4).Value = "EstablishmentYear";
            sheet.Cell(1, 5).Value = "LeagueId";
            sheet.Cell(1, 6).Value = "StadiumId";

            int row = 2;
            foreach (var team in teams)
            {
                sheet.Cell(row, 1).Value = team.TeamId;
                sheet.Cell(row, 2).Value = team.Name;
                sheet.Cell(row, 3).Value = team.City;
                sheet.Cell(row, 4).Value = team.EstablishmentYear;
                sheet.Cell(row, 5).Value = team.LeagueId;
                sheet.Cell(row, 6).Value = team.StadiumId;
                row++;
            }
        }

        static void ExportStadiums(XLWorkbook workbook, List<Stadium> stadiums)
        {
            var sheet = workbook.AddWorksheet("Stadiums");
            sheet.Cell(1, 1).Value = "StadiumId";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Capacity";
            sheet.Cell(1, 4).Value = "City";
            sheet.Cell(1, 5).Value = "Address";

            int row = 2;
            foreach (var stadium in stadiums)
            {
                sheet.Cell(row, 1).Value = stadium.StadiumId;
                sheet.Cell(row, 2).Value = stadium.Name;
                sheet.Cell(row, 3).Value = stadium.Capacity;
                sheet.Cell(row, 4).Value = stadium.City;
                sheet.Cell(row, 5).Value = stadium.Address;
                row++;
            }
        }
    }

    class Season
    {
        public int SeasonId { get; set; }
        public string Name { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }

    class League
    {
        public int LeagueId { get; set; }
        public string Name { get; set; }
    }

    class Team
    {
        public int TeamId { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public int EstablishmentYear { get; set; }
        public int LeagueId { get; set; }
        public string StadiumId { get; set; }
    }

    class Stadium
    {
        public string StadiumId { get; set; }
        public string Name { get; set; }
        public int Capacity { get; set; }
        public string City { get; set; }
        public string Address { get; set; }
    }

    class Player
    {
        public int PlayerId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Nationality { get; set; }
        public DateTime BirthDate { get; set; }
        public int Height { get; set; }
        public string PositionId { get; set; }
    }
    class Position
    {
        public string PositionId { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
    }

    class Coach
    {
        public string CoachId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Nationality { get; set; }
        public string ExperienceLevel { get; set; }
        public int CoachingLicenses { get; set; }
        public string CoachingStyle { get; set; }
        public string TrainingMethods { get; set; }
        public string PlayerDevelopmentFocus { get; set; }
        public string InjuryManagement { get; set; }
        public string TeamSelectionCriteria { get; set; }
        public string CommunicationStyle { get; set; }
        public string MotivationalTechniques { get; set; }
        public DateTime BirthDate { get; set; }
    }

    class Referee
    {
        public string RefereeId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Nationality { get; set; }
        public string Email { get; set; }
        public string PhoneNumber { get; set; }
        public DateTime BirthDate { get; set; }
    }


    class CoachAssignment
    {
        public string AssignmentId { get; set; }
        public string CoachId { get; set; }
        public int TeamId { get; set; }
        public int SeasonId { get; set; }
    }
    class PlayerAssignment
    {
        public string AssignmentId { get; set; }
        public int PlayerId { get; set; }
        public int TeamId { get; set; }
        public int SeasonId { get; set; }


    }


    class Match
    {
        public string MatchId { get; set; }
        public int HomeTeamId { get; set; }
        public int AwayTeamId { get; set; }
        public string StadiumId { get; set; }
        public string RefereeId { get; set; }
        public DateTime Date { get; set; }
        public int Attendance { get; set; }
        public int SeasonId { get; set; }
        public int? WinnerTeamId { get; set; } // Null if draw
        public int LeagueId { get; set; }
    }

    class Goal
    {
        public int GoalId { get; set; }
        public string MatchId { get; set; }
        public int PlayerId { get; set; }
        //public int AssisterId { get; set; }
        public int TeamId { get; set; }
        public int SeasonId { get; set; }
        public int LeagueId { get; set; }
    }




}
