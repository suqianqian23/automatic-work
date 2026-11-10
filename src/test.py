DEFAULT_SETTING: dict = {
    "process": ["技术变更通知单", "封面", "变更履历", "PBOM清单", ""],
}

value = DEFAULT_SETTING.get("process")
print(value)
print(type(value))

for i in value:
    print(value)
