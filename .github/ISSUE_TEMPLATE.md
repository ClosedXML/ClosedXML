## Read and complete the full issue template

Do not randomly delete sections. They are here for a reason.

**Do you want to request a *feature* or report a *bug*?**
- [x] Bug
- [ ] Feature
- [ ] Question

**Did you test against the latest CI build?**
- [ ] Yes
- [ ] No

If you answered `No`, please test with [the latest development build](https://ci.appveyor.com/project/ClosedXML/ClosedXML/branch/develop/artifacts) first.

**Version of ClosedXML**

e.g. 0.95.3

**What is the current behavior?**

Complete this.

**What is the expected behavior or new feature?**

Complete this.

**Is this a regression from the previous version?**

Regressions get higher priority. Test against the latest build of the previous minor version. For example, if you experience a problem on v0.95.3, check whether it the problem occurred in v0.94.2 too. 

## Reproducibility
**This is an important section. Read it carefully. Failure to do so will cause a 'RTFM' comment.**

Without a code sample, it is unlikely that your issue will get attention. Don't be lazy. Do the effort and assist the developers to reproduce your problem. Code samples should be [minimal complete and verifiable](https://stackoverflow.com/help/mcve). Sample spreadsheets should be attached whenever applicable. Remove sensitive information.

**Code to reproduce problem:**
```c#
public void Main()
{
    // Code standards:
    // - Fully runnable. I should be able to copy and paste this code into a 
    //   console application and run it without having to edit it much.
    // - Declare all your variables (this follows from the previous point)
    // - The code should be a minimal code sample to illustrate issue. The code 
    //   samples on the wiki are good examples of the terseness that I want. Don't
    //   post your full application.
}
```

- [ ] I attached a sample spreadsheet.  (You can drag files on to this issue)
