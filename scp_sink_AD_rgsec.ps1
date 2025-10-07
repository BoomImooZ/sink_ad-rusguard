
cls
$users = @{}
$serverName = "localhost"  # или IP-адрес
$databaseName = "RusGuardDB"
$connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=true;"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$connection.Open()
$command = $connection.CreateCommand() # Не запоминай в ней свою команду - каждый раз эта переменная используется под разные команды!!
$Def_Group_SQL_for_user = 'GUID_DEF_GROUP_ON_SQL_RGSEC' # Группа пользователей в которую будут добавлены пользователи на SQL


#Укажи сюда нужные группы, можно как списком так и 1 группу
$GroupList ="SKUD_1floor","SKUD_office_5floor"#,"2" #Введи названия групп или передай в эту переменную нужные группы


#Создание группы доступа
function Sql-ADD_AcsAccessLevel {
    param(
        [Parameter(Mandatory=$true)]
        $Group
    )
    $reader.Close()
    $command.CommandText = ""
    $command.Parameters.Clear()
    $command.CommandText = "SELECT * FROM dbo.AcsAccessLevel WHERE Name = @Name"
    $command.Parameters.Add("@Name", [System.Data.SqlDbType]::NVarChar, 255).Value = $Group
    $reader = $command.ExecuteReader()
    
    if (-not $reader.Read()){
        $reader.Close()
        $command.CommandText = ""
        $command.Parameters.Clear()
        $command.CommandText = "INSERT INTO dbo.AcsAccessLevel (Name, IsRemoved, NumberOfAccessPoints) VALUES ('"+$Group +"', 0, 0)"
        "Создана группа: "+$group+" В количестве "+$command.ExecuteNonQuery()
        
    }
    $reader.Close()
}


#Добавление фото пользователю
function Sql-ADD_EmployeePhoto {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    $user_id = ((Sql-GetUser $User)[$User['samaccountname']])['_id']

    #Проверка есть ли фото?
    $reader.Close()
    $command.Parameters.Clear()
    $command.CommandText = "SELECT * FROM dbo.EmployeePhoto WHERE EmployeeID = @EmployeeID AND PhotoNumber = @PhotoNumber"
    $command.Parameters.Add("@EmployeeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]$User_id
    $command.Parameters.Add("@PhotoNumber", [System.Data.SqlDbType]::Int).Value = 1
    $reader=$command.ExecuteReader()

    if (-not $reader.read()){
        #добавление нового фото
        $reader.Close()
        $command.Parameters.Clear()
        $command.CommandText = "INSERT INTO dbo.EmployeePhoto (EmployeeID, PhotoNumber, Photo, EmployeeImageType) VALUES (@EmployeeID, @PhotoNumber, @Photo, @EmployeeImageType)"
        $command.Parameters.Add("@EmployeeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]$User_id
        $command.Parameters.Add("@PhotoNumber", [System.Data.SqlDbType]::Int).Value = 1
        $command.Parameters.Add("@Photo", [System.Data.SqlDbType]::VarBinary, -1).Value = $User.photo
        $command.Parameters.Add("@EmployeeImageType", [System.Data.SqlDbType]::Int).Value = 0
        "Затронуто строк: "+ $command.ExecuteNonQuery()
        "Добавлено фото польователю "+$User.SamAccountName
    }Else{
           "Удалите фото пользователя "+$User.SamAccountName+" с сервера вручную и тогда добавьте снова."
    }
    $reader.Close()
}


#Добавление пользователей в группы доступа
function Sql-EmplAcsAccessLevel {
    param(
        [Parameter(Mandatory=$true)]
        $User_id,
        $Group_id
    )

        #проверка - выдана ли пользователю группа, если не выдана, то выдать
        $reader.Close()
        $command.CommandText = ""
        $command.Parameters.Clear()
        $command.CommandText = "SELECT * FROM dbo.EmployeeAcsAccessLevel WHERE EmployeeID = @EmployeeID AND AcsAccessLevelID = @AcsAccessLevelID"
        $command.Parameters.Add("@EmployeeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = $User_id
        $command.Parameters.Add("@AcsAccessLevelID", [System.Data.SqlDbType]::UniqueIdentifier).Value = $Group_id
        $reader = $command.ExecuteReader()
        
        if ( -not $reader.Read()){
            $reader.Close()
            $command.CommandText = ""
            $command.Parameters.Clear()
            $command.CommandText = "INSERT INTO dbo.EmployeeAcsAccessLevel (EmployeeID, AcsAccessLevelID) VALUES (@EmployeeID, @AcsAccessLevelID)"
            $command.Parameters.Add("@EmployeeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = $User_id
            $command.Parameters.Add("@AcsAccessLevelID", [System.Data.SqlDbType]::UniqueIdentifier).Value = $Group_id
            #"Пользоователь с _id "+ $User_id +" будет добавлен в группу доступа: " + $Group_id
            "Число затронутых строк: "+$command.ExecuteNonQuery()
            $result = "Пользоователь с _id "+ $User_id +" будет добавлен в группу доступа: " + $Group_id
        }
        $reader.Close()
        return $result
}


#Берет инфу из AD о пользователях в группе безопасности
function Get-GroupInfo {
    param(
        [Parameter(Mandatory=$true)]
        $Groups
    )
    $result= @{}
    foreach($group in $Groups){
        $GroupMembers = Get-ADGroupMember -Identity $group | Where-Object { $_.objectClass -eq 'User' } | Select-Object samaccountname, name
        #Write-Output "Группа: $Group"
        foreach($member in $GroupMembers){
            $users[$member.samaccountname] = @{
                samaccountname = $member.samaccountname
                firstname = ($member.name -split(" "))[1]     #Имя
                secondname = ($member.name -split(" "))[2]    #Отчество
                lastname = ($member.name -split(" "))[0]      #Фамилия
                title = (Get-ADUser -Identity $member.samaccountname -Properties title | Select-Object title).title
                photo = (Get-ADUser -Identity $member.samaccountname -Properties jpegphoto | Select-Object jpegphoto).jpegphoto[0]
            }
        $result[$group] = $users
        }
    }
    return $result
}


#Вывод данных о существующем пользователе, если пусто - то пользователь не существует
function Sql-GetUser {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    $userData = @{}
    $sqluser = @{}
    $command.CommandText = ""
    $command.CommandText = "SELECT * FROM dbo.Employee WHERE (FirstName = '"+$User['firstname']+"') And (LastName = '"+$User['lastname']+"') And (SecondName = '"+$User['secondname']+"') And (IsRemoved = '0') And (Comment= '"+$User["samaccountname"]+"')"
    $reader = $command.ExecuteReader()
    while ($reader.Read()) {
        
        for ($i=0; $i -lt $reader.FieldCount; $i++){
            $columnName = $reader.GetName($i)
            $columnValue = $reader[$i]
            $userData[$columnName] = $columnValue
            
        }        
        $sqluser[$User.samaccountname] = $userData        
    }
    $reader.Close()
    return $sqluser
    
}


#Вывод данных о группах доступа
function Sql-GetAcsAccessLevel {

    $AcsAccessLevel_data = @{}
    $Data = @{}


    foreach ($Group in $GroupList){
        $reader.Close()
        $command.CommandText = ""
        $command.Parameters.Clear()
        $command.CommandText = "SELECT * FROM dbo.AcsAccessLevel WHERE Name = @Name"
        $command.Parameters.Add("@Name", [System.Data.SqlDbType]::NVarChar, 255).Value = $Group
        $reader = $command.ExecuteReader()
        if (-not $reader.Read()){
            $reader.Close() 
            Sql-ADD_AcsAccessLevel -Group $Group
        }
        $reader.Close()
    }

    $reader.Close()
    $command.CommandText = ""
    $command.CommandText = "SELECT * FROM dbo.AcsAccessLevel"
    $reader = $command.ExecuteReader()
    while ($reader.Read()) {

        for ($i=0; $i -lt $reader.FieldCount; $i++){
            $columnName = $reader.GetName($i)
            $columnValue = $reader[$i]
            $Data[$columnName] = $columnValue
            
        }
        $AcsAccessLevel_data[$data['Name']] = $Data
        $data = @{}
        
    }
    $reader.Close()

    return $AcsAccessLevel_data
    
}


#Создает пользователя в SQL
function Sql-AddUser {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    $reader.Close()
    $command.CommandText = ""
    $command.Parameters.Clear()
    $command.CommandText = "SELECT * FROM dbo.Employee WHERE (IsRemoved = '1') And (Comment= '"+$User["samaccountname"]+"')"
    $reader = $command.ExecuteReader()
    if ($reader.Read()){
        $reader.Close()
        $command.CommandText = "UPDATE dbo.Employee SET IsRemoved = 0 WHERE Comment = '"+$User["samaccountname"]+"'"
        "Число затронутых строк: "+$command.ExecuteNonQuery()
        $result = "Пользователь "+$User["samaccountname"]+" восстановлен"
    }else{
        $reader.Close()
        $command.CommandText = ""
        $command.Parameters.Clear()
        $command.CommandText = "INSERT INTO dbo.Employee (FirstName, SecondName, LastName, EmployeeGroupID, Comment, IsAccessLevelsInherited) VALUES (@FirstName, @SecondName, @LastName, @EmployeeGroupID, @Comment, @IsAccessLevelsInherited)"
        $null = $command.Parameters.AddWithValue("@FirstName", $User.firstname)
        $null = $command.Parameters.AddWithValue("@SecondName", $User.secondname)
        $null = $command.Parameters.AddWithValue("@LastName", $User.lastname)
        $null = $command.Parameters.AddWithValue("@Comment", $User.samaccountname)
        $null = $command.Parameters.AddWithValue("@IsAccessLevelsInherited", 0)
        $null = $command.Parameters.AddWithValue("@EmployeeGroupID", $Def_Group_SQL_for_user)
        "Число затронутых строк: "+$command.ExecuteNonQuery()
        $result = "Пользователь "+$User["samaccountname"]+" создан"
        $reader.Close()
    }
    Sql-ADD_EmployeePhoto -User $User
    Sql-ADD_AcsKeys -User $User
    return $result

}


#Обертка над добавлением пользователя в группы доступа
function Sql-ADD_EmployeeAcsAccessLevel {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )
    foreach ($group in $Users.Keys){
        $group_id = (Sql-GetAcsAccessLevel)[$Group]['_id']
        foreach ($user in $Users[$group].keys){
            #Если нет пользователя или отключен - то создаст/включит его

            if ((-not (Get-SQLuser -User $Users[$group][$user])[$Users[$group][$user]['samaccountname']] ) -or
                (((Get-SQLuser -User $Users[$group][$user])[$Users[$group][$user]['samaccountname']]).IsRemoved)) {

                Sql-AddUser $Users[$group][$user]
            }
            $user_id = ((Get-SQLuser -User $Users[$group][$user])[$Users[$group][$user]['samaccountname']])['_id']
            Sql-EmplAcsAccessLevel -User_id $user_id -Group_id $group_id
        }
    }
    return "Закончил работу по группам"
}


#Добавление пользователю карты
function Sql-ADD_AcsKeys {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    $User
    $counter_rows = 0
    $sqldata =@{}
    $rows = @{}
    $user_id = ((Sql-GetUser $User)[$User['samaccountname']])['_id']
    
    $reader.Close()
    $command.CommandText = ""
    $command.CommandText = "SELECT * FROM dbo.AcsKeys"
    $reader = $command.ExecuteReader()
    while ($reader.Read()) {
        $counter_rows+= 1
        for ($i=0; $i -lt $reader.FieldCount; $i++){
            $columnName = $reader.GetName($i)
            $columnValue = $reader[$i]
            $sqldata[$columnName] = $columnValue
        }    
        $rows[$counter_rows] = $sqldata
    }
    $card_id = $rows.Count+1
    $reader.Close()
    $command.CommandText = ""
    $command.Parameters.Clear()
    $command.CommandText = "SELECT * FROM dbo.AcsKeys WHERE Name = @Name AND CardTypeID = @CardTypeID"    
    $command.Parameters.Add("@Name", [System.Data.SqlDbType]::NVarChar, 255).Value = $User['samaccountname']
    $command.Parameters.Add("@CardTypeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]"8B239530-6592-4B6B-A473-5FABCE7B4F82"
    $reader = $command.ExecuteReader()
    if (-not $reader.Read()){
        $reader.Close()
        $command.CommandText = ""
        $command.Parameters.Clear()
        $command.CommandText = "INSERT INTO dbo.AcsKeys (KeyNumber, Name, CardTypeID, StartDate) VALUES (@KeyNumber, @Name, @CardTypeID, @StartDate)"    
        $command.Parameters.Add("@KeyNumber", [System.Data.SqlDbType]::BigInt).Value = $card_id
        $command.Parameters.Add("@Name", [System.Data.SqlDbType]::NVarChar, 255).Value = $User['samaccountname']
        $command.Parameters.Add("@CardTypeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]"8B239530-6592-4B6B-A473-5FABCE7B4F82"
        $command.Parameters.Add("@StartDate", [System.Data.SqlDbType]::Date).Value = (Get-Date -Format yyyy-MM-dd)
        "Число созданных карт: "+$command.ExecuteNonQuery()
        #"карта пользователя "+$User["samaccountname"]+" создана"
        $reader.Close()
        $command.CommandText = ""
        $command.Parameters.Clear()
        $command.CommandText = "INSERT INTO dbo.AcsKey2EmployeeAssignment (AcsKeyID, EmployeeId, AssignmentModificationDateTime, AssignmentModificationType, IndexNumberOfAssignedKey) VALUES (@AcsKeyID, @EmployeeId, @AssignmentModificationDateTime, @AssignmentModificationType, @IndexNumberOfAssignedKey)"    
        $command.Parameters.Add("@AcsKeyID", [System.Data.SqlDbType]::BigInt).Value = $card_id
        $command.Parameters.Add("@EmployeeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]$User_id
        $command.Parameters.Add("@AssignmentModificationDateTime", [System.Data.SqlDbType]::datetime2).Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        $command.Parameters.Add("@AssignmentModificationType", [System.Data.SqlDbType]::int).Value = 0
        $command.Parameters.Add("@IndexNumberOfAssignedKey", [System.Data.SqlDbType]::int).Value = 1
        "Число связанных карт с пользователем: "+$command.ExecuteNonQuery()
        #"карта свзяана с пользователем "+$User["samaccountname"]+" создана"
        $reader.Close()
        $result = "Карта "+$card_id+" Связана с пользователем "+$User.SamAccountName
    }

    return $result
}




#Нужен скрипт который уберет из группы доступа, если отозвали права
#Удаление из группы доступа
function Sql-DEL_EmplAcsAccessLevel {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )

        $user_ids= @{}
        $all_user_ids = @{}
        foreach ($group in $Users.keys){
            $group_id=[guid]((Sql-GetAcsAccessLevel)[$group])['_id']
            $reader.Close()
            $command.CommandText = ""
            $command.Parameters.Clear()
            $command.CommandText = "SELECT EmployeeID FROM dbo.EmployeeAcsAccessLevel WHERE AcsAccessLevelID = @AcsAccessLevelID"
            $command.Parameters.Add("@AcsAccessLevelID", [System.Data.SqlDbType]::UniqueIdentifier).Value = $group_id
            $reader = $command.ExecuteReader()
            while ($reader.read()){
                $user_ids[$reader[0]] = ""
            }
            foreach ($user in ($Users[$group]).keys){
                $reader.Close()
                $command.CommandText = ""
                $command.Parameters.Clear()
                $command.CommandText = "SELECT _id FROM dbo.Employee WHERE Comment = @Comment"
                $command.Parameters.Add("@Comment", [System.Data.SqlDbType]::nvarchar, -1).Value = ($user)
                $reader = $command.ExecuteReader()
                while ($reader.read()){
                    $all_user_ids[$reader[0]] = ""
                }
            }
            foreach ($u_acl in $user_ids.keys){
                $find = 0
                foreach ($u_adg in $all_user_ids.keys){
                    if ($u_acl -eq $u_adg){
                        $find = 1
                    }
                }
                if ($find -eq 0){
                    $reader.Close()
                    $command.CommandText = ""
                    $command.Parameters.Clear()
                    $command.CommandText = "DELETE FROM dbo.EmployeeAcsAccessLevel where EmployeeID = @EmployeeID AND AcsAccessLevelID = @AcsAccessLevelID";
                    $command.Parameters.Add("@EmployeeID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]$u_acl
                    $command.Parameters.Add("@AcsAccessLevelID", [System.Data.SqlDbType]::UniqueIdentifier).Value = [guid]$group_id
                    "Удалил доступ к группам, у пользователей: "+$command.ExecuteNonQuery()
                    $reader.Close()

                }
                $find = 0
            }
        }
        
        $reader.Close()
        return "Закончил забирать доступы у пользователей"
}

# Вызов Тут мейн:

#Основной рабочий скрипт 
$users = Get-GroupInfo -Groups $GroupList
Sql-ADD_EmployeeAcsAccessLevel  -Users $users
Sql-DEL_EmplAcsAccessLevel -Users $users
