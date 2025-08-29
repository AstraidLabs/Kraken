#if UNIT_TESTS
using System;
using Xunit;

public class SppHelperTests
{
    [Fact]
    public void GetKmsClientInfo_ReturnsExpectedValues()
    {
        // Arrange – use a known SKU GUID that is KMS‑enabled
        Guid skuId = new Guid("00000000-0000-0000-0000-000000000000");
        // Act
        var infos = SppHelper.GetSkuActivationInfo(skuId);
        // Assert
        Assert.NotEmpty(infos);
        Assert.All(infos, info =>
        {
            Assert.False(string.IsNullOrEmpty(info.CustomerPID));
            Assert.False(string.IsNullOrEmpty(info.KeyManagementServiceName));
        });
    }
}
#endif
