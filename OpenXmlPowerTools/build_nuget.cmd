nuget pack -IncludeReferencedProjects -Build -Prop Configuration=Release -Prop Platform=AnyCPU 
xcopy /Q /Y *.nupkg "../../../NuGet/*"