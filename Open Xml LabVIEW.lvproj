﻿<?xml version='1.0' encoding='UTF-8'?>
<Project Type="Project" LVVersion="21008000">
	<Property Name="NI.LV.All.SaveVersion" Type="Str">21.0</Property>
	<Property Name="NI.LV.All.SourceOnly" Type="Bool">true</Property>
	<Property Name="NI.Project.Description" Type="Str"></Property>
	<Item Name="My Computer" Type="My Computer">
		<Property Name="server.app.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.control.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.tcp.enabled" Type="Bool">false</Property>
		<Property Name="server.tcp.port" Type="Int">0</Property>
		<Property Name="server.tcp.serviceName" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.tcp.serviceName.default" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.vi.callsEnabled" Type="Bool">true</Property>
		<Property Name="server.vi.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="specify.custom.address" Type="Bool">false</Property>
		<Item Name="Tests" Type="Folder">
			<Item Name="memory leaks.vi" Type="VI" URL="../Test Open Xml/memory leaks.vi"/>
			<Item Name="Test Open Xml.lvclass" Type="LVClass" URL="../Test Open Xml/Test Open Xml.lvclass"/>
		</Item>
		<Item Name="Excel Cell Address to Numeric Indices.vi" Type="VI" URL="../Source/Excel Cell Address to Numeric Indices.vi"/>
		<Item Name="Numeric Indices to Excel Cell Address.vi" Type="VI" URL="../Source/Numeric Indices to Excel Cell Address.vi"/>
		<Item Name="Open Xml.lvlib" Type="Library" URL="../Source/Open Xml.lvlib"/>
		<Item Name="Dependencies" Type="Dependencies"/>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
