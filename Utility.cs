<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <!-- TargetFramework / Nullable / ImplicitUsings 由 Directory.Build.props 统一管理 -->
    <IsPackable>false</IsPackable>
    <!-- 本项目不引用真 Core / exe csproj(避免其未完成的 Core 引用),而是直接链接 exe 逻辑源 + Fake.Core。 -->
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
    <PackageReference Include="SharpZipLib" Version="1.4.2" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.0" />
  </ItemGroup>

  <!-- 被测源:链接 exe 的逻辑文件(不含 Program.cs,避免第二入口与额外依赖)。 -->
  <ItemGroup>
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\PhotoImportJob.cs"     Link="Sut\PhotoImportJob.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\PhotoImportOptions.cs" Link="Sut\PhotoImportOptions.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\WatermarkStore.cs"       Link="Sut\WatermarkStore.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\SingleInstanceLock.cs"   Link="Sut\SingleInstanceLock.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\RunSummary.cs"           Link="Sut\RunSummary.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\AppliedManifestStore.cs" Link="Sut\AppliedManifestStore.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\XmlPhotoReader.cs"        Link="Sut\XmlPhotoReader.cs" />
  </ItemGroup>

  <ItemGroup>
    <None Include="..\COD.FirmwideDirectory.PhotoImportTool\mock_photo.xml"
          Link="TestData\mock_photo.xml"
          CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>

</Project>
