import openpyxl
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Dict, List

@dataclass
class IPConfig:
    original_name: str
    role: str  # 'master' or 'slave'
    read_write: str
    original_bit_width: int
    original_frequency: int
    original_clk_domain: str
    final_bit_width: int = 0
    final_frequency: int = 0
    final_protocol: str = "-"
    final_clk_domain: str = "-"
    connected_interconnect: str = "-"

@dataclass
class InterconnectConfig:
    name: str
    bit_width: int
    frequency: int
    protocol: str
    clk_domain: str
    master_ips: List[str]  # Original IP names (ip1, ip2, etc.)
    slave_ips: List[str]   # Original IP names (ip3, ip4, etc.)
    connected_interconnects: List[str]

class ConfigGenerator:
    def __init__(self):
        self.ip_configs: Dict[str, IPConfig] = {}  # Key: M1/S1 format
        self.original_ip_map: Dict[str, str] = {}  # Original name to M1/S1 mapping
        self.interconnect_configs: Dict[str, InterconnectConfig] = {}
        self.script_dir = Path(__file__).parent
        self.master_count = 0
        self.slave_count = 0
    
    def find_excel_file(self) -> Path:
        """Finds the first .xls or .xlsx file in script directory"""
        for file in self.script_dir.glob("*.*"):
            if file.suffix.lower() in ('.xls', '.xlsx'):
                return file
        raise FileNotFoundError("No Excel file found in script directory")
    
    def read_excel(self) -> None:
        """Reads the Excel file with two sheets (IPs and Interconnects)"""
        excel_file = self.find_excel_file()
        print(f"\n[DEBUG] Found Excel file: {excel_file.name}")
        
        try:
            workbook = openpyxl.load_workbook(excel_file)
            
            # =============================================
            # Process IP sheet (first sheet)
            # =============================================
            print("\n[DEBUG] Processing IP Sheet:")
            ip_sheet = workbook.worksheets[0]
            for row_idx, row in enumerate(ip_sheet.iter_rows(min_row=2, values_only=True), start=2):
                if not row[0]:  # Skip if IP name is empty
                    continue
                
                original_name = str(row[0]).strip()
                
                # Determine if master or slave
                is_master = bool(row[5])  # Master ID column
                role = 'master' if is_master else 'slave'
                
                # Generate M1, M2 or S1, S2 names
                if is_master:
                    self.master_count += 1
                    ip_name = f"M{self.master_count}"
                else:
                    self.slave_count += 1
                    ip_name = f"S{self.slave_count}"
                
                # Store mapping from original name to M1/S1 name
                self.original_ip_map[original_name] = ip_name
                
                ip_config = IPConfig(
                    original_name=original_name,
                    role=role,
                    read_write=str(row[1]),
                    original_bit_width=int(row[2]),
                    original_frequency=int(row[3]),
                    original_clk_domain=str(row[4]),
                    final_bit_width=int(row[2]),  # Initialize with original values
                    final_frequency=int(row[3]),
                    final_clk_domain=str(row[4])
                )
                self.ip_configs[ip_name] = ip_config
                print(f"[DEBUG] Created IP {ip_name} (originally {original_name}): {ip_config}")
            
            # =============================================
            # Process Interconnect sheet (second sheet)
            # =============================================
            if len(workbook.worksheets) > 1:
                print("\n[DEBUG] Processing Interconnect Sheet:")
                interconnect_sheet = workbook.worksheets[1]
                for row_idx, row in enumerate(interconnect_sheet.iter_rows(min_row=2, values_only=True), start=2):
                    if not row[0]:  # Skip if interconnect name is empty
                        continue
                    
                    # Clean and split comma-separated lists
                    master_ips = [ip.strip() for ip in str(row[5]).split(',')] if row[5] else []
                    slave_ips = [ip.strip() for ip in str(row[6]).split(',')] if row[6] else []
                    connected_ics = [ic.strip() for ic in str(row[7]).split(',')] if row[7] else []
                    
                    interconnect = InterconnectConfig(
                        name=str(row[0]),
                        bit_width=int(row[1]),
                        frequency=int(row[2]),
                        protocol=str(row[3]),
                        clk_domain=str(row[4]),
                        master_ips=master_ips,
                        slave_ips=slave_ips,
                        connected_interconnects=connected_ics
                    )
                    self.interconnect_configs[interconnect.name] = interconnect
                    print(f"[DEBUG] Created Interconnect {interconnect.name}: {interconnect}")
            
            # =============================================
            # Apply interconnect properties to IPs
            # =============================================
            self._apply_interconnect_properties()
            
            # Debug print final IP configurations
            print("\n[DEBUG] Final IP Configurations:")
            for ip_name, config in self.ip_configs.items():
                print(f"{ip_name}: {config}")
            
            print(f"\nProcessed {len(self.ip_configs)} IPs and {len(self.interconnect_configs)} interconnects")
            
        except Exception as e:
            print(f"\n[ERROR] Reading Excel file: {str(e)}")
            raise
    
    def _apply_interconnect_properties(self):
        """Updates IP properties based on connected interconnects"""
        print("\n[DEBUG] Applying interconnect properties:")
        
        # First create reverse mapping from original IP names to interconnects
        ip_to_interconnect = {}
        
        for ic_name, ic_config in self.interconnect_configs.items():
            # Process masters
            for original_ip in ic_config.master_ips:
                if original_ip in self.original_ip_map:
                    mapped_name = self.original_ip_map[original_ip]
                    ip_to_interconnect[mapped_name] = ic_name
                    print(f"[DEBUG] Mapped master {original_ip} ({mapped_name}) to {ic_name}")
                else:
                    print(f"[WARNING] Master IP {original_ip} not found in IP sheet")
            
            # Process slaves
            for original_ip in ic_config.slave_ips:
                if original_ip in self.original_ip_map:
                    mapped_name = self.original_ip_map[original_ip]
                    ip_to_interconnect[mapped_name] = ic_name
                    print(f"[DEBUG] Mapped slave {original_ip} ({mapped_name}) to {ic_name}")
                else:
                    print(f"[WARNING] Slave IP {original_ip} not found in IP sheet")
        
        # Now update IP configurations
        for ip_name, interconnect_name in ip_to_interconnect.items():
            if ip_name in self.ip_configs and interconnect_name in self.interconnect_configs:
                ip_config = self.ip_configs[ip_name]
                ic_config = self.interconnect_configs[interconnect_name]
                
                ip_config.connected_interconnect = interconnect_name
                ip_config.final_bit_width = ic_config.bit_width
                ip_config.final_frequency = ic_config.frequency
                ip_config.final_protocol = ic_config.protocol
                ip_config.final_clk_domain = ic_config.clk_domain
                
                print(f"[DEBUG] Updated {ip_name} with {interconnect_name} properties")
    
    def generate_config_file(self) -> None:
        """Generates config.txt in the same directory"""
        output_file = self.script_dir / "config.txt"
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                # Write header
                header = [
                    "IP NAME".ljust(15),
                    "TYPE".ljust(10),
                    "READ/WRITE".ljust(15),
                    "BIT WIDTH".ljust(15),
                    "FREQUENCY".ljust(15),
                    "PROTOCOL".ljust(15),
                    "CLK DOMAIN".ljust(15),
                    "INTERCONNECT".ljust(15),
                    "ORIGINAL IP".ljust(15)
                ]
                f.write("".join(header) + "\n")
                f.write("=" * 150 + "\n")
                
                # Write IP configurations
                for ip_name, config in sorted(self.ip_configs.items()):
                    row = [
                        ip_name.ljust(15),
                        ("MASTER" if config.role == 'master' else "SLAVE").ljust(10),
                        config.read_write.ljust(15),
                        str(config.final_bit_width).ljust(15),
                        str(config.final_frequency).ljust(15),
                        config.final_protocol.ljust(15),
                        config.final_clk_domain.ljust(15),
                        config.connected_interconnect.ljust(15),
                        config.original_name.ljust(15)
                    ]
                    f.write("".join(row) + "\n")
            
            print(f"\nConfig file generated: {output_file.name}")
            
            # Print final output to console
            print("\nFinal Output Preview:")
            print("".join(header))
            print("-" * 150)
            for ip_name, config in sorted(self.ip_configs.items()):
                row = [
                    ip_name.ljust(15),
                    ("MASTER" if config.role == 'master' else "SLAVE").ljust(10),
                    config.read_write.ljust(15),
                    str(config.final_bit_width).ljust(15),
                    str(config.final_frequency).ljust(15),
                    config.final_protocol.ljust(15),
                    config.final_clk_domain.ljust(15),
                    config.connected_interconnect.ljust(15),
                    config.original_name.ljust(15)
                ]
                print("".join(row))
            
        except Exception as e:
            print(f"\n[ERROR] Generating config file: {str(e)}")
            raise

if __name__ == "__main__":
    print("Auto Excel to Config Generator")
    print("=" * 40)
    print("Looking for Excel file in script directory...")
    
    generator = ConfigGenerator()
    try:
        generator.read_excel()
        generator.generate_config_file()
        print("\nDone! Check config.txt in the same folder")
    except Exception as e:
        print(f"\nError occurred: {str(e)}")