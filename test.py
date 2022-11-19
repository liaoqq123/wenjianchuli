import transitions

# class Matter():
#     states = ['solid', 'liquid', 'gas', 'plasma']  # 状态有固态、液态、气态、等离子态
#     transitions = [
#         {'trigger': 'melt', 'source': 'solid', 'dest': 'liquid'},
#         {'trigger': 'evaporate', 'source': 'liquid', 'dest': 'gas'},
#         {'trigger': 'sublimate', 'source': 'solid', 'dest': 'gas'},
#         {'trigger': 'ionize', 'source': 'gas', 'dest': 'plasma'}
#     ]
#
#     def __init__(self):
#         self.machine = Machine(model=self, states=Matter.states, transitions=Matter.transitions, initial='solid')
#
#
# lump = Matter()
# print(lump.state)  # solid
# lump.melt()
# print(lump.state)  # liquid

# 连接协议状态机
from transitions.extensions import HierarchicalMachine as Machine
from transitions.extensions.nesting import NestedState
class excel_inspect(object):
    pass

excels = excel_inspect()


states = [
    'int',
    'str',
    'intgroup',
    'strgroup'
]

transition = [
    ['a', '*', 'int'],
    ['b', '*', 'str'],
    ['c', '*', 'intgroup'],
    ['d', '*', 'strgroup']
]


mach = transitions.Machine(
    model=excels,
    states=states,
    transitions=transition,
    initial='int'
)

aa = input('你随便输入一个')
if aa in ['a', 'b', 'c', 'd']:
    print(excels.aa())
else:
    print('别TM瞎输')