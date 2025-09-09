---
name: Generic Issue
about: generic issue template
title: ''
labels: 
assignees: ''

---

### Problem:
    title-hint:= cpt[ModuleName_(bas|frm)] - 'object not set'
    description-hint:= be detailed
    bonus-points: include screenshot(s)

```vba
'even-more-bonus-points: offending code snippet here
```

### Solution:
    how should this issue be addressed?

### Todo:
- [ ] assign this issue
- [ ] assign issue type (Bug, Feature, Task)
- [ ] label the issue with codemodule (so fixes can be aggregated and hotfixed together)
- [ ] checkout appropriate branch and create `topic` branch
- [ ] design, code and test
- [ ] update codemodule `<cpt_version>x.y.z</cpt_version>`
- [ ] update CurrentVersions.xml **manually**
- [ ] commit message should begin with `Issue #XXX - `
- [ ] commit the change(s)
- [ ] merge `topic` into appropriate branch(es) and push
- [ ] delete topic branch
