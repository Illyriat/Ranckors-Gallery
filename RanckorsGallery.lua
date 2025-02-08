-- Instead of calling require or dofile, just get the globals.
local Strings = _G["RanckorsGalleryStrings"]
local Colors  = _G["RanckorsGalleryColors"]

local RanckorsGallery = {}
local addonName = "RanckorsGallery"
local savedVars

function RanckorsGallery:OnAddonLoaded(event, addon)
    if addon ~= addonName then return end

    EVENT_MANAGER:UnregisterForEvent(addonName, EVENT_ADD_ON_LOADED)
    savedVars = ZO_SavedVars:NewAccountWide("RanckorsGallerySavedVars", 1, nil, {})
    self:InitializeEvents()
    self:InitializeUI()
    d(Strings.ADDON_LOADED)
end

function RanckorsGallery:InitializeEvents()
    -- Your event registration here.
end

function RanckorsGallery:InitializeUI()
    -- Your UI initialization code here.
end

function RanckorsGallery:OnPlayerActivated(event)
    d("Welcome to Ranckors Gallery!")
    -- You can add any additional actions to perform upon login here.
end

function RanckorsGallery:OnFurniturePlaced(event, furnitureData)
    d(Colors.Green .. "Furniture placed: " .. tostring(furnitureData.itemId) .. Colors.Reset)
end

EVENT_MANAGER:RegisterForEvent(addonName, EVENT_ADD_ON_LOADED, function(event, addon)
    RanckorsGallery:OnAddonLoaded(event, addon)
end)