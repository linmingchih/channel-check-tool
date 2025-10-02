import re
from collections import defaultdict
from pyedb import Edb


#gnds = ['DGND', ]
gnds = ['GND', ]

src_path = r'E:/ANSYS_2025/2025_0120_CCT_test/F5Z/a787140a_DFA-0627-1650-v172.brd'
#src_path = r"E:/ANSYS_2025/2025_0120_CCT_test/F5Z/a787140a_DFA-0627-1650-v172.aedb"
edb_path = src_path.replace('.brd', '.aedb')
#new_edb_path = src_path.replace('.aedb', '_new.aedb')

edb = Edb(src_path, edbversion='2024.2', )
#edb.nets.plot()
print(edb.core_siwave.pin_groups)
#%%
cct_comps = ['U43','U42']
comp_nets = defaultdict(list)
#pattern = '(\w+)_(DATA|DQ|DM|DQS_H|DQS_L|DQSP|DQSN)_?(\d+)'
pattern = '(\w+)_(DATA|DQ|DM|DQS_H|DQS_L|DQSP|DQSN)_?(\d+|<\d+>)'   #DDR4_DQ<17>
pattern = 'DDR_D\d+|DDR_DQS\d+_P|DDR_DQS\d+_N|DDR_DMI\d+'    #DDR_D7
delete_list = []
all_net_names = []

for net_name, net_obj in edb.nets.nets.items():
    if net_name in gnds:
        continue
    
    if not set(net_obj.components.keys()) <= set(cct_comps):
        continue
    
    m = re.search(pattern, net_name, re.IGNORECASE)
    if m:
        all_net_names.append(net_name)
        for comp in net_obj.components:
            comp_nets[comp].append(net_name)
            
    else:
        delete_list.append(net_name)
#%%

edb.cutout(all_net_names, gnds, extent_type="Bounding")

#edb.nets.delete(delete_list)

ports = []
pg_pair = []
for comp, nets in comp_nets.items():
  
    #gn = edb.core_siwave.create_pin_group_on_net(comp, 'DGND', f'{comp}_gnd')
    gn = edb.core_siwave.create_pin_group_on_net(comp, 'GND', f'{comp}_gnd')
    tn = gn[1].create_port_terminal(50)
    for net in nets:
        gp = edb.core_siwave.create_pin_group_on_net(comp, net,  f'{comp}_{net}')
        tp = gp[1].create_port_terminal(50)
        port_name = f'port;{comp};{net}'
        ports.append(port_name)
        tp.SetName(port_name)
        tp.SetReferenceTerminal(tn)
        
frequency_sweep=[["linear count", "0", "1kHz", 1],
                 ["log scale", "1kHz", "0.1GHz", 10],
                 ["linear scale", "0.1GHz", "10GHz", "0.1GHz"]]

setup = edb.create_siwave_syz_setup()
setup.add_frequency_sweep('mysetup', frequency_sweep)
#edb.save_edb_as(new_edb_path)
edb.save_edb_as(edb_path)
edb.close_edb()

#%%

from pyaedt import Hfss3dLayout
#hfss = Hfss3dLayout(new_edb_path, version='2025.1', remove_lock=True, non_graphical=False)
hfss = Hfss3dLayout(edb_path, version='2024.2', remove_lock=True, non_graphical=False)

hfss.design_settings['SParamExport'] = True

hfss.save_project()
hfss.analyze(cores=12)
hfss.release_desktop()


