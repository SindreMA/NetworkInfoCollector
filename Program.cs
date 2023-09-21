using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Management;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NetworkInfoCollector
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var PSVMs = GetPsObjectList("Get-VM");
            List<object> VMs = new List<object>();
            foreach (var item in PSVMs)
            {
                var l = item.BaseObject;
                VMs.Add(l);
            }
            List<object> DNSs = new List<object>();
            var DnsClientServerAddress = GetPsObjectList("Get-DnsClientServerAddress");
            foreach (var item in DnsClientServerAddress)
            {
                var InterfaceIndex = item.Properties["InterfaceIndex"].Value;
                var DNSAdresses = item.Properties["ServerAddresses"].Value;
                var InterfaceAlias = item.Properties["InterfaceAlias"].Value;
                var ElementName = item.Properties["ElementName"].Value;
                DNSs.Add(new { InterfaceAlias, InterfaceIndex , ElementName, DNSAdresses });
            }


            var sd = GetPsObjectList("Get-NetIPConfiguration");
            foreach (var item in sd)
            {
                var AllIPAddresses = item.Properties["AllIPAddresses"].Value;
                var DNSServer = item.Properties["DNSServer"].Value;
                var InterfaceAlias = item.Properties["InterfaceAlias"].Value;
                var IPv4Address = item.Properties["IPv4Address"].Value;
                var IPv6Address = item.Properties["IPv6Address"].Value;
                var IPv6LinkLocalAddress = item.Properties["IPv6LinkLocalAddress"].Value;
                var get_IPv4DefaultGateway = item.Properties["get_IPv4DefaultGateway"].Value;
                var get_IPv6DefaultGateway = item.Properties["get_IPv6DefaultGateway"].Value;
            }



            var PSNicTeams = GetPsObjectList("get-netlbfoteam");
            var PSNicTeamMembers = GetPsObjectList("get-NetLbfoTeamMember");
            var PSNicIPAdresses = GetPsObjectList("Get-NetIPAddress");

            var NicIPAdresses = new List<object>();
            foreach (var item in PSNicIPAdresses)
            {
                var InterfaceAlias = item.Properties["InterfaceAlias"].Value;
                var IPAddress = item.Properties["IPAddress"].Value;
                var InterfaceIndex = item.Properties["InterfaceIndex"].Value;
                NicIPAdresses.Add(new {InterfaceAlias, InterfaceIndex,IPAddress });
            }

            var NicTeamMembers = new List<NicTeamMember>();
            foreach (var item in PSNicTeamMembers)
            {
                var Team = item.Properties["Team"].Value;
                var Name = item.Properties["Name"].Value;
                var InstanceID = item.Properties["InstanceID"].Value;
                var InterfaceDescription = item.Properties["InterfaceDescription"].Value;
                var EnabledState = item.Properties["EnabledState"].Value;
                var EnabledDefault = item.Properties["EnabledDefault"].Value;
                var RequestedState = item.Properties["RequestedState"].Value;
                var OperationalMode = item.Properties["OperationalMode"].Value;
                var TransmitLinkSpeed = item.Properties["TransmitLinkSpeed"].Value;
                var ReceiveLinkSpeed = item.Properties["ReceiveLinkSpeed"].Value;
                var TransitioningToState = item.Properties["TransitioningToState"].Value;
                var AdministrativeMode = item.Properties["AdministrativeMode"].Value;
                var FailureReason = item.Properties["FailureReason"].Value;


                NicTeamMembers.Add(new NicTeamMember {
                    NetworkAdapterInfo = new Networkadapterinfo(),
                    InstanceID = InstanceID as string,
                    Name = Name as string,
                    Team = Team as string,
                    EnabledDefault = EnabledDefault as object,
                    EnabledState = EnabledState as object,
                    RequestedState = RequestedState as object,
                    OperationalMode = OperationalMode as object,
                    TransmitLinkSpeed = TransmitLinkSpeed as object,
                    ReceiveLinkSpeed = ReceiveLinkSpeed as object,
                    TransitioningToState = TransitioningToState as object,
                    AdministrativeMode = AdministrativeMode as object,
                    FailureReason = FailureReason as object });
            }
            var NicTeams = new List<NicTeam>();
            foreach (var item in PSNicTeams)
            {
                var InstanceID = item.Properties["InstanceID"].Value;
                var Name = item.Properties["Name"].Value;
                var LoadBalancingAlgorithm = item.Properties["LoadBalancingAlgorithm"].Value;
                var Status = item.Properties["Status"].Value;
                var TeamingMode = item.Properties["TeamingMode"].Value;

                NicTeams.Add(new NicTeam {
                    Members =  new List<NicTeamMember>(),
                    InstanceID = InstanceID as string,
                    Name = Name as string,
                    LoadBalancingAlgorithm = LoadBalancingAlgorithm as object,
                    Status = Status as object,
                    TeamingMode = TeamingMode as object });
            }

            
            var NetAdapters = new List<Networkadapterinfo>();
            foreach (var item in GetPsObjectList("get-netadapter"))
            {
                var DeviceID = item.Properties["DeviceID"].Value;
                var DriverDescription = item.Properties["DriverDescription"].Value;
                var DriverProvider = item.Properties["DriverProvider"].Value;
                var DriverVersionString = item.Properties["DriverVersionString"].Value;

                var InterfaceDescription = item.Properties["InterfaceDescription"].Value;
                var InterfaceName = item.Properties["InterfaceName"].Value;

                var Name = item.Properties["Name"].Value;
                var NetworkAddresses = item.Properties["NetworkAddresses"].Value;
                var MacAddress = item.Properties["PermanentAddress"].Value;
                var Speed = item.Properties["Speed"].Value;
                var SystemName = item.Properties["SystemName"].Value;

                Networkadapterinfo NetAdapter = new Networkadapterinfo {
                    DeviceID = DeviceID as string,
                    DriverDescription = DriverDescription as string,
                    DriverProvider = DriverProvider as string,
                    DriverVersionString = DriverVersionString as string,
                    InterfaceName = InterfaceName as string,
                    InterfaceDescription = InterfaceDescription as string,
                    Name = Name as string,
                    NetworkAddresses = NetworkAddresses as string[],
                    MacAddress = MacAddress as string,
                    Speed = Speed as object,
                    SystemName = SystemName as string };
                NetAdapters.Add(NetAdapter);
            }

            foreach (var team in NicTeams)
            {
                Console.WriteLine(team.Name);
                team.Members = NicTeamMembers.Where(x => x.Team == team.Name).ToList();
                foreach (var member in team.Members)
                {
                    Console.WriteLine("-"+member.Name);
                    if (NetAdapters.Exists(x => x.Name == member.Name))
                    {
                        
                        var NetworkAdapterInfo = NetAdapters.Find(x => x.Name == member.Name);
                        member.NetworkAdapterInfo = NetworkAdapterInfo;
                        Console.WriteLine("--" + NetworkAdapterInfo.MacAddress);
                    }


                }
            }

            var json = JsonConvert.SerializeObject(new { NicTeams = NicTeams, NetAdapters,NicIPAdresses, VMs  }, Formatting.Indented ,
                        new JsonSerializerSettings()
                        {
                            ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                            
                        });

            File.WriteAllText("Result.json",json);
        }


        public static List<PSObject> GetPsObjectList( string Item, List<string> parameter = null)
        {
            parameter = parameter ?? new List<string>();
            var ls = new List<PSObject>();

                using (var psRunSpace = GetRunspaceMsol())
                {
                    using (PowerShell powershell = PowerShell.Create())
                    {
                        powershell.Runspace = psRunSpace;
                        Command createGroupCmd = new Command(Item);
                        foreach (var item in parameter)
                        {
                            if (item.Contains(" "))
                            {
                                var split = item.Split(' ');
                                CommandParameter userParam = new CommandParameter(split[0], split[1]);
                                createGroupCmd.Parameters.Add(userParam);
                            }
                            else
                            {
                                createGroupCmd.Parameters.Add(item);
                            }
                        }
                        powershell.Commands.AddCommand(createGroupCmd);
                        var res = powershell.Invoke();
                        var errors2 = powershell.Streams.Error.ReadAll();
                        ls.AddRange(res);
                        powershell.Commands.Clear();
                    }
                    psRunSpace.Close();
                }
          
            return ls;
        }

        public static Runspace GetRunspaceMsol()
        {
            Runspace psRunSpace = null;
            try
            {
                InitialSessionState initialSession = InitialSessionState.CreateDefault();
                initialSession.ImportPSModule(new[] { "Hyper-V" });
                psRunSpace = RunspaceFactory.CreateRunspace(initialSession);
                psRunSpace.Open();
            }
            catch
            {
            }
            return psRunSpace;
        }
        public class NicTeam
        {
            public List<NicTeamMember> Members { get; set; }
            public string InstanceID { get; set; }
            public string Name { get; set; }
            public object LoadBalancingAlgorithm { get; set; }
            public object Status { get; set; }
            public object TeamingMode { get; set; }
        }
        public class NicTeamMember
        {
            public Networkadapterinfo NetworkAdapterInfo { get; set; }
            public string InstanceID { get; set; }
            public string Name { get; set; }
            public string Team { get; set; }
            public object EnabledDefault { get; set; }
            public object EnabledState { get; set; }
            public object RequestedState { get; set; }
            public object OperationalMode { get; set; }
            public object TransmitLinkSpeed { get; set; }
            public object ReceiveLinkSpeed { get; set; }
            public object TransitioningToState { get; set; }
            public object AdministrativeMode { get; set; }
            public object FailureReason { get; set; }
        }
        public class Networkadapterinfo
        {
            public string DeviceID { get; set; }
            public string DriverDescription { get; set; }
            public string DriverProvider { get; set; }
            public string DriverVersionString { get; set; }
            public string InterfaceName { get; set; }
            public string InterfaceDescription { get; set; }
            public string Name { get; set; }
            public string[] NetworkAddresses { get; set; }
            public string MacAddress { get; set; }
            public object Speed { get; set; }
            public string SystemName { get; set; }
            public object IP { get; set; }
        }


    }
}




















