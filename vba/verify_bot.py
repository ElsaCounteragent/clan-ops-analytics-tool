import os
import logging
import discord
from discord.ext import commands
import aiohttp
import urllib.parse

# ============================================================
# Configuration & Logging
# ============================================================
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Securely load configuration from environment variables
DISCORD_BOT_TOKEN = os.getenv('DISCORD_BOT_TOKEN')
COC_API_KEY = os.getenv('COC_API_KEY')
VERIFIED_ROLE_ID = int(os.getenv('VERIFIED_ROLE_ID', '0'))

if not all([DISCORD_BOT_TOKEN, COC_API_KEY, VERIFIED_ROLE_ID]):
    logging.critical("Missing required environment variables. Please check your .env file.")
    exit(1)

# ============================================================
# Discord UI Components
# ============================================================
class VerificationModal(discord.ui.Modal, title='Verify your Clash of Clans account'):
    player_tag = discord.ui.TextInput(
        label='What is your player tag?',
        style=discord.TextStyle.short,
        placeholder='Example: #VP0QG80Y',
        required=True
    )
    api_token = discord.ui.TextInput(
        label='What is your API token?',
        style=discord.TextStyle.short,
        placeholder='Example: pd6smtn9',
        required=True
    )

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        
        tag = self.player_tag.value.strip().upper()
        token = self.api_token.value.strip()
        
        if not tag.startswith('#'):
            tag = '#' + tag
            
        safe_tag = urllib.parse.quote(tag)
        
        headers = {
            'Authorization': f'Bearer {COC_API_KEY}',
            'Accept': 'application/json'
        }

        async with aiohttp.ClientSession() as session:
            try:
                # 1. Verify the token provided by the user
                verify_url = f'https://api.clashofclans.com/v1/players/{safe_tag}/verifytoken'
                async with session.post(verify_url, headers=headers, json={"token": token}) as resp:
                    
                    if resp.status != 200:
                        if resp.status == 403:
                            logging.error("403 Forbidden - Supercell API Key IP whitelist is likely outdated.")
                            await interaction.followup.send(
                                "Verification failed. The bot's API Key IP whitelist might be incorrect.", 
                                ephemeral=True
                            )
                        else:
                            error_text = await resp.text()
                            logging.error(f"API Status {resp.status}: {error_text}")
                            await interaction.followup.send(f"Verification failed. Unexpected API error ({resp.status}).", ephemeral=True)
                        return

                    verify_data = await resp.json()
                    
                    if verify_data.get('status', '').lower() == 'ok':
                        # 2. Fetch player profile
                        profile_url = f'https://api.clashofclans.com/v1/players/{safe_tag}'
                        async with session.get(profile_url, headers=headers) as prof_resp:
                            if prof_resp.status == 200:
                                prof_data = await prof_resp.json()
                                in_game_name = prof_data.get('name', 'Unknown')[:32] # Discord 32-char limit
                                
                                await self.update_discord_profile(interaction, in_game_name)
                            else:
                                await interaction.followup.send("Token valid, but failed to fetch profile data.", ephemeral=True)
                    else:
                        await interaction.followup.send("Verification failed. Please check your Tag and Token.", ephemeral=True)

            except aiohttp.ClientError as e:
                logging.error(f"Network Error: {e}")
                await interaction.followup.send("Failed to connect to the API. Please try again later.", ephemeral=True)
            except Exception as e:
                logging.exception("Unexpected error during verification")
                await interaction.followup.send("An unexpected error occurred.", ephemeral=True)

    async def update_discord_profile(self, interaction: discord.Interaction, in_game_name: str):
        """Handles idempotent role and nickname updates."""
        if not isinstance(interaction.user, discord.Member) or not interaction.guild:
            await interaction.followup.send("This command must be used in a server.", ephemeral=True)
            return

        role = interaction.guild.get_role(VERIFIED_ROLE_ID)
        if not role:
            await interaction.followup.send("Error: Verified Role not found in server settings.", ephemeral=True)
            return

        try:
            has_role = role in interaction.user.roles
            name_matches = interaction.user.nick == in_game_name or (interaction.user.nick is None and interaction.user.name == in_game_name)

            if has_role and name_matches:
                await interaction.followup.send(f"You are already verified and up to date, {in_game_name}!", ephemeral=True)
            elif has_role:
                await interaction.user.edit(nick=in_game_name)
                await interaction.followup.send(f"Welcome back, {in_game_name}. Your name has been updated.", ephemeral=True)
            else:
                await interaction.user.edit(nick=in_game_name, roles=interaction.user.roles + [role])
                await interaction.followup.send(f"Success! Welcome, {in_game_name}.", ephemeral=True)
                
        except discord.Forbidden:
            await interaction.followup.send("I verified you, but my role is too low to change your nickname/roles.", ephemeral=True)
        except discord.HTTPException as e:
            logging.error(f"Discord HTTP Error: {e}")
            await interaction.followup.send(f"Verified, but failed to update Discord profile: {e}", ephemeral=True)


class VerificationView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="Verify Account", style=discord.ButtonStyle.green, custom_id="verify_button_1")
    async def verify_btn(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_modal(VerificationModal())

    @discord.ui.button(label="Where is my Tag?", style=discord.ButtonStyle.secondary, custom_id="info_tag_btn")
    async def info_tag_btn(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(
            "**Finding your Player Tag:**\n1. Open Clash of Clans.\n2. Tap your profile level.\n3. Your tag is below your name.", 
            ephemeral=True
        )

    @discord.ui.button(label="Where is my API Token?", style=discord.ButtonStyle.secondary, custom_id="info_token_btn")
    async def info_token_btn(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(
            "**Finding your API Token:**\n1. Open Settings.\n2. Tap 'More Settings'.\n3. Scroll to 'API Token' and tap 'Show'.", 
            ephemeral=True
        )

# ============================================================
# Bot Setup
# ============================================================
intents = discord.Intents.default()
intents.message_content = True
intents.members = True

class MyBot(commands.Bot):
    def __init__(self):
        super().__init__(command_prefix="!", intents=intents)

    async def setup_hook(self):
        self.add_view(VerificationView())

    async def on_ready(self):
        logging.info(f'Logged in securely as {self.user}')

bot = MyBot()

@bot.command()
@commands.has_permissions(administrator=True)
async def spawn_verify(ctx):
    await ctx.send(">>> 🛡️ Click The Green Button to verify your Clash of Clans account 🛡️ <<<", view=VerificationView())

if __name__ == "__main__":
    bot.run(DISCORD_BOT_TOKEN)